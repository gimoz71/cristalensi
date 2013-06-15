<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_clsImageSize.asp"-->
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
                            <h3>Categoria scelta: lampade moderne</h3>
                            <p>
                                <i>Stai cercando lampade moderne per l'illuminazione della tua casa? il catalogo dei prodotti esposti è molto ampio, così abbiamo diviso gli articoli per illuminazione moderna in base alla loro destinazione: scegli tra lampade a sospensione e lampadari moderni, plafoniere moderne, lampade a parete e applique moderni, da tavolo, lumini e abat-jour moderne, lampade da ufficio e da scrivania, piantane moderne e lampade da terra. Tutti prodotti con uno stile moderno e di design dalle più importanti marche e produttori.</i>.
                            </p>
                            <ul class="produttori clearfix">
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano wfwefewfe</span>
                                    </a>
                                </li>
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano</span>
                                    </a>
                                </li>
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano</span>
                                    </a>
                                </li>
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano</span>
                                    </a>
                                </li>
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano</span>
                                    </a>
                                </li>
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano</span>
                                    </a>
                                </li>
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano</span>
                                    </a>
                                </li>
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano</span>
                                    </a>
                                </li>
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano</span>
                                    </a>
                                </li>
                                <li>
                                    <a href="#">
                                        <img src="images/example.jpg">
                                        <span class="button_link">Applique di vetro murano</span>
                                    </a>
                                </li>
                                
                            </ul>
                            <h4 class="area">Produttori</h4>
                            <p>Sei interessato ad una specifica marca? Ricerca il tuo prodotto tramite la nostra selezione di produttori
                            </p>
                            <select name="FkProduttore" id="FkProduttore" class="form" onChange="invia_produttore()">
                                <option value="0">Seleziona un produttore</option>

                                <option value="20">Alta Tensione</option>

                                <option value="48">Antonangeli</option>

                                <option value="40">Arte Luce</option>

                                <option value="47">Artemide</option>

                                <option value="1">Atom</option>

                                <option value="39">Augenti Illuminazione</option>

                                <option value="12">Belfiore</option>

                                <option value="21">CRISTALENSI</option>

                                <option value="31">EGLO</option>

                                <option value="8">Ellequattro</option>

                                <option value="4">Eurokeramic</option>

                                <option value="60">FARO</option>

                                <option value="27">FB Braga</option>

                                <option value="26">Flami</option>

                                <option value="46">Flos</option>

                                <option value="11">Fustilamp</option>

                                <option value="6">Garden Luce</option>

                                <option value="59">GEA LUCE</option>

                                <option value="44">Genex</option>

                                <option value="28">Geol</option>

                                <option value="13">Gibas</option>

                                <option value="45">Globo Lighting</option>

                                <option value="25">HOMEGA </option>

                                <option value="17">Ideal Lux</option>

                                <option value="43">I-LèD</option>

                                <option value="5">Illuminando</option>

                                <option value="53">IMAS</option>

                                <option value="51">ISMOS</option>

                                <option value="9">Isoluce</option>

                                <option value="37">ITALAMP</option>

                                <option value="22">Italexport</option>

                                <option value="23">LAMPADINE</option>

                                <option value="55">LineaLight</option>

                                <option value="58">Livos</option>

                                <option value="49">Lucifero</option>

                                <option value="3">Lux</option>

                                <option value="29">Mantoan Luce</option>

                                <option value="24">MC LUCE</option>

                                <option value="50">Microluce</option>

                                <option value="56">MURANO</option>

                                <option value="16">Murano Luce</option>

                                <option value="19">Pan International</option>

                                <option value="18">Perenz</option>

                                <option value="7">Scamm</option>

                                <option value="14">Selene Illuminazione</option>

                                <option value="33">Sforzin</option>

                                <option value="57">Sovil </option>

                                <option value="32">Surya</option>

                                <option value="10">Top Light Illuminazione</option>

                                <option value="52">TOSCOT</option>

                                <option value="54">Tràddel</option>

                                <option value="42">Vesoi</option>

                                <option value="30">Vistosi</option>

                            </select>
                            <p>Oppure consulta direttamente la pagina con <a href="#">l'elenco completo dei produttori</a></p>
                        </div>
                    </div>
                </div>
                <!--#include file="inc_sx_prodotti.asp"-->
            </div>
        </div>
         <!--#include file="inc_footer.asp"-->
    </body>
</html>
