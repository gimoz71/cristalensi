<!--#include file="inc_strConn.asp"-->
<!doctype html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Cristalensi</title>
        <!--[if (lt IE 9)&(!IEMobile)]>
        <link rel="stylesheet" type="text/css" href="enhanced.css" />
        <![endif]-->
        <!--[if lt IE 9]>
        <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
        <script src="js/media-queries-ie.js"></script>
        <![endif]-->
        <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
        <script src="js/init.js"></script>
        <link href="css/css.css" rel="stylesheet" type="text/css">
        <style type="text/css">
            .clearfix:after {
                content: ".";
                display: block;
                height: 0;
                clear: both;
                visibility: hidden;
            }
        </style>
    </head>
    <body>
        <div id="wrap">
            <!--#include file="inc_header.asp"-->
            <div id="main-content">
                
                <div id="content-sidebar-wrap" >
                    <div id="content">
                        <div>
                        
                        	<h3 style="font-size: 14px; display: inline; border: none;">Autenticazione cliente</h3>
                            <p>Se sei già iscritto, e quindi hai già Login (Email) e Password, non è necessario che ti iscriva nuovamente, è sufficiente inserire i dati di accesso qui sotto e sarai riconosciuto immediatamente.</p>
							<div class="iscrizione clearfix">
                            	<div class="table">
                                    <div class="tr">
                                        <div class="td">
                                            Login <input name="nominativo" type="text" class="form" id="nominativo" size="30" maxlength="50" value="">
                                        </div>
                                        <div class="td">
                                            Password <input name="Rag_Soc" type="text" class="form" id="Rag_Soc" size="30" maxlength="50" value=""><input type="submit" value="accedi">
                                            <p style="text-align: center"><a href="#">Clicca quì per recuperare la password</a></p>
                                        </div>
                                       
                                    </div>
                                </div>
                            </div>
                            <hr>
                            <h3 style="font-size: 14px; display: inline; border: none;">Registrazione cliente</h3>
                            <p>In questa pagina puoi inserire i tuoi dati per registrarti a Cristalensi.<br />
                                Informazione importante: è necessario che l'indirizzo Email sia un'indirizzo funzionante e che usi normalmente, in quanto ti verranno spedite comunicazioni relativamente agli ordini e ai prodotti.<br />
                                Ti ricordiamo inoltre che l'indirizzo Email lo dovrai utilizzare come Login per accedere ai tuoi futuri ordini.
                             </p>
                            <div class="iscrizione clearfix">
                                
                                <div class="table">
                                    <form method="post" action="iscrizione.asp?mode=1&amp;pkid=0" name="newsform" onSubmit="return verifica();">
                                    </form>
                                    <div class="tr">
                                        <div class="td">
	                                        Nome e Cognome (*)<br>
                                            <input name="nominativo" type="text" class="form" id="nominativo" size="30" maxlength="50" value="">
                                        </div>
                                        <div class="td">
                                        	Ragione sociale ( nel caso in cui si tratti di un'Azienda )<br>
                                            <input name="Rag_Soc" type="text" class="form" id="Rag_Soc" size="30" maxlength="50" value="">
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        Codice Fiscale<br>
                                            <input name="cod_fisc" type="text" class="form" id="cod_fisc" size="20" maxlength="16" value="">
                                        </div>
                                        <div class="td">
                                        Partita IVA ( nel caso in cui si tratti di un'Azienda )<br>
                                            <input name="PartitaIVA" type="text" class="form" id="PartitaIVA" size="20" maxlength="11" value="">
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        	Indirizzo (*)<br>
                                            <input name="indirizzo" type="text" class="form" id="indirizzo" size="30" maxlength="100" value="">
                                        </div>
                                        <div class="td">
                                        	CAP<br>
                                            <input name="cap" type="text" class="form" id="cap" size="7" maxlength="5" value="">
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
	                                        Città (*)<br>
                                            <input name="citta" type="text" class="form" id="citta" size="30" maxlength="50" value="">
                                        </div>
                                        <div class="td">
	                                        Provincia<br>
                                            <input type="text" name="provincia" id="provincia" value="" size="3" maxlength="2" class="form">
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>Nazione</div>
                                    </div>
                                    <div class="tr">
                                        Italia:&nbsp;&nbsp;Si&nbsp;
                                            <input type="radio" name="italia" value="Sì" checked="">
                                            &nbsp;&nbsp;No&nbsp;
                                            <input type="radio" name="italia" value="No">
                                            &nbsp;Altra nazione
                                            <input name="nazionediversa" type="text" class="form" id="nazionediversa" size="30" maxlength="50" value="">
                                        
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        Telefono (*)<br>
                                            <input name="telefono" type="text" class="form" id="telefono" size="30" maxlength="50" value="">
                                        </div>
                                        <div class="td">
	                                        Fax<br>
                                            <input name="fax" type="text" class="form" id="fax" size="30" maxlength="50" value="">
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div class="td">
                                        	<strong>E-mail</strong> (*) - Verrà usata come <strong>Login</strong> per i futuri ordini<br>
                                            <input name="email" type="text" class="form" id="email" size="30" maxlength="100" value="">
                                        </div>
                                        <div class="td">
                                        <strong>Password</strong> (*)<br>
                                            <input name="password" type="password" class="form" id="pw" size="30" maxlength="50" value="">
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div >Autorizzazione a ricevere Email </div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <input type="radio" name="aut_email" value="True" checked="">
                                            &nbsp;Si&nbsp;&nbsp;
                                            <input type="radio" name="aut_email" value="False">
                                            &nbsp;No
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>&nbsp;</div>
                                    </div>
                                    <div class="tr">
                                        <div>Condizioni sul trattamento dei dati personali</div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <textarea name="privacy" cols="80" rows="5" readonly class="form">INFORMAZIONI RELATIVE AL TRATTAMENTO DI DATI PERSONALI Ai sensi dell'art. 10 della L. n°675 del 31/12/1996, l'Azienda informa l'interessato che i dati che lo riguardano, forniti dall'interessato medesimo, formeranno oggetto di trattamento nel rispetto della normativa sopra richiamata. Tali dati verranno trattati per finalita' gestionali, commerciali, promozionali. Il conferimento dei dati alla nostra Azienda e' assolutamente facoltativo. I dati acquisiti potranno essere comunicati e diffusi in osservanza di quanto disposto all'articolo 20 della legge 675/96 allo scopo di perseguire le finalita' sopra indicate. Il titolare del trattamento e' Cristalensi s.n.c. con sede in via arti e mestieri, 1 Montelupo F.no (FI) , ove e' altresì domiciliato il responsabile protempore del trattamento, i cui dati identificativi possono essere acquisiti presso il Registro pubblico tenuto dal Garante, o presso la sede legale dell'Azienda. L'Azienda informa altresì l'Interessato che questi potra' esercitare i diritti previsti dall'articolo 13 della legge 675/96, ossia: Conoscere gratuitamente, mediante accesso al Registro Generale del Garante, l'esistenza di trattamenti di dati che possono riguardarlo; Ottenere da Cristalensi s.n.c., - con un contributo spese solo in caso di risposta negativa - la conferma dell'esistenza o meno nei propri archivi di dati che lo riguardino, ed avere la loro comunicazione e l'indicazione della logica e delle finalita' su cui si basa il trattamento. La richiesta e' rinnovabile dopo novanta giorni; Ottenere la cancellazione, la trasformazione in forma anonima ed il blocco dei dati trattati in violazione di legge; Ottenere l'aggiornamento, la rettifica o l'integrazione dei dati; Ottenere l'attestazione che la cancellazione, l'aggiornamento, la rettifica o l'integrazione siano portate a conoscenza di coloro che abbiano avuto comunicazione dei dati; Opporsi gratuitamente al trattamento dei dati che lo riguardano. </textarea>
                                        </div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <input name="chekka" type="checkbox" onClick="accetta(this)">
                                            Accetta le condizioni </div>
                                    </div>
                                    <div class="tr">
                                        <div>&nbsp;</div>
                                    </div>
                                    <div class="tr">
                                        <div>
                                            <input name="Submit" type="submit" class="form" value="Salva" align="absmiddle" disabled="">
                                            &nbsp;
                                            <input name="Submit2" type="reset" class="form" value="Annulla">
                                            (*) campo obbligatorio </div>
                                    </div>
                                    <div class="tr">
                                        <div>&nbsp;</div>
                                    </div>
                                </div>

                            
                                <table style="display: none;">
                                    <form method="post" action="iscrizione.asp?mode=1&amp;pkid=0" name="newsform" onSubmit="return verifica();">
                                    </form>
                                    <tbody>
                                        <tr align="left">
                                            <td>Nome e Cognome (*)</td>
                                            <td>Ragione sociale ( nel caso in cui si tratti di un'Azienda ) </td>
                                        </tr>
                                        <tr align="left">
                                            <td ><input name="nominativo" type="text" class="form" id="nominativo" size="30" maxlength="50" value=""></td>
                                            <td><input name="Rag_Soc" type="text" class="form" id="Rag_Soc" size="30" maxlength="50" value=""></td>
                                        </tr>
                                        <tr align="left">
                                            <td>Codice Fiscale</td>
                                            <td>Partita IVA ( nel caso in cui si tratti di un'Azienda ) </td>
                                        </tr>
                                        <tr align="left">
                                            <td><input name="cod_fisc" type="text" class="form" id="cod_fisc" size="20" maxlength="16" value=""></td>
                                            <td><input name="PartitaIVA" type="text" class="form" id="PartitaIVA" size="20" maxlength="11" value=""></td>
                                        </tr>
                                        <tr align="left">
                                            <td>Indirizzo (*)</td>
                                            <td>Cap</td>
                                        </tr>
                                        <tr align="left">
                                            <td><input name="indirizzo" type="text" class="form" id="indirizzo" size="30" maxlength="100" value=""></td>
                                            <td><input name="cap" type="text" class="form" id="cap" size="7" maxlength="5" value=""></td>
                                        </tr>
                                        <tr align="left">
                                            <td>Città (*)</td>
                                            <td>Provincia</td>
                                        </tr>
                                        <tr align="left">
                                            <td><input name="citta" type="text" class="form" id="citta" size="30" maxlength="50" value=""></td>
                                            <td><input type="text" name="provincia" id="provincia" value="" size="3" maxlength="2" class="form"></td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2">Nazione</td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2">Italia:&nbsp;&nbsp;Si&nbsp;
                                                <input type="radio" name="italia" value="Sì" checked="">
                                                &nbsp;&nbsp;No&nbsp;
                                                <input type="radio" name="italia" value="No">
                                                &nbsp;Altra nazione
                                                <input name="nazionediversa" type="text" class="form" id="nazionediversa" size="30" maxlength="50" value=""></td>
                                        </tr>
                                        <tr align="left">
                                            <td>Telefono (*)</td>
                                            <td>Fax</td>
                                        </tr>
                                        <tr align="left">
                                            <td><input name="telefono" type="text" class="form" id="telefono" size="30" maxlength="50" value=""></td>
                                            <td><input name="fax" type="text" class="form" id="fax" size="30" maxlength="50" value=""></td>
                                        </tr>
                                        <tr align="left">
                                            <td><strong>E-mail</strong> (*) - Verrà usata come <strong>Login</strong> per i futuri ordini </td>
                                            <td><strong>Password</strong> (*)</td>
                                        </tr>
                                        <tr align="left">
                                            <td><input name="email" type="text" class="form" id="email" size="30" maxlength="100" value=""></td>
                                            <td><input name="password" type="password" class="form" id="pw" size="30" maxlength="50" value=""></td>
                                        </tr>
                                       
                                        <tr align="left">
                                            <td colspan="2">Autorizzazione a ricevere Email </td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2"><input type="radio" name="aut_email" value="True" checked="">
                                                &nbsp;Si&nbsp;&nbsp;
                                                <input type="radio" name="aut_email" value="False">
                                                &nbsp;No</td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2">&nbsp;</td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2">Condizioni sul trattamento dei dati personali</td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2"><textarea name="privacy" cols="80" rows="5" readonly class="form">INFORMAZIONI RELATIVE AL TRATTAMENTO DI DATI PERSONALI Ai sensi dell'art. 10 della L. n°675 del 31/12/1996, l'Azienda informa l'interessato che i dati che lo riguardano, forniti dall'interessato medesimo, formeranno oggetto di trattamento nel rispetto della normativa sopra richiamata. Tali dati verranno trattati per finalita' gestionali, commerciali, promozionali. Il conferimento dei dati alla nostra Azienda e' assolutamente facoltativo. I dati acquisiti potranno essere comunicati e diffusi in osservanza di quanto disposto all'articolo 20 della legge 675/96 allo scopo di perseguire le finalita' sopra indicate. Il titolare del trattamento e' Cristalensi s.n.c. con sede in via arti e mestieri, 1 Montelupo F.no (FI) , ove e' altresì domiciliato il responsabile protempore del trattamento, i cui dati identificativi possono essere acquisiti presso il Registro pubblico tenuto dal Garante, o presso la sede legale dell'Azienda. L'Azienda informa altresì l'Interessato che questi potra' esercitare i diritti previsti dall'articolo 13 della legge 675/96, ossia: Conoscere gratuitamente, mediante accesso al Registro Generale del Garante, l'esistenza di trattamenti di dati che possono riguardarlo; Ottenere da Cristalensi s.n.c., - con un contributo spese solo in caso di risposta negativa - la conferma dell'esistenza o meno nei propri archivi di dati che lo riguardino, ed avere la loro comunicazione e l'indicazione della logica e delle finalita' su cui si basa il trattamento. La richiesta e' rinnovabile dopo novanta giorni; Ottenere la cancellazione, la trasformazione in forma anonima ed il blocco dei dati trattati in violazione di legge; Ottenere l'aggiornamento, la rettifica o l'integrazione dei dati; Ottenere l'attestazione che la cancellazione, l'aggiornamento, la rettifica o l'integrazione siano portate a conoscenza di coloro che abbiano avuto comunicazione dei dati; Opporsi gratuitamente al trattamento dei dati che lo riguardano. </textarea></td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2"><input name="chekka" type="checkbox" onClick="accetta(this)">
                                                Accetta le condizioni
                                            </td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2">&nbsp;</td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2"><input name="Submit" type="submit" class="form" value="Salva" align="absmiddle" disabled="">
                                                &nbsp;
                                                <input name="Submit2" type="reset" class="form" value="Annulla">
                                                (*) campo obbligatorio
                                            </td>
                                        </tr>
                                        <tr align="left">
                                            <td colspan="2">&nbsp;</td>
                                        </tr>
                                    </tbody>
                                </table>
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