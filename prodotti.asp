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
                            <div class="slogan">
                                <h3>Eccezionale sconto!!! Nessun costo di spedizione per ordini superiori a 250€</h3>
                                <p>Per ordini inferiori a 250€ il costo di spedizione è di 10€.<br> Condizioni valide solo per le spedizioni in tutta Italia, isole comprese.</p>
                            </div>
                            <h3>Prodotti in vetrina</h3>
                            <p>
                                <i>Questa è una breve selezione di prodotti che rappresentano la nostra galleria.<br>
                                Per consultare tutto il catalogo ed accedere ai singoli prodotti, potete scegliere una categoria sulla sinistra.<br> 
                                Ogni prodotto ha una propria scheda dettagliata, per accederci è sufficiente cliccare sul nome o sulla foto del prodotto</i>.
                            </p>
                            <ul class="prodotti clearfix">
                                <li class="clearfix">
                                    <div class="thumb">
                                    <a href="#">
                                        <img src="images/example.jpg">
                                    </a>
                                    </div>
                                    <div class="data">
                                        <a href="scheda.html"><strong>Applique di vetro murano</strong></a> <span class="produttore">produttore: <a href="#"><strong>Illumnando</strong></a></span>
                                        <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vel mauris vitae urna ornare convallis a eu elit. Nulla varius lobortis lorem a molestie...</p>
                                        <a class="button_link" href="scheda.html">Scheda prodotto</a>
                                        <p class="cart clearfix"><span class="price">Prezzo listino: <span>155€</span></span> <span class="cristalprice">Prezzo listino: 155€</span><a href="#" class="cart-link button_link"><span>Inserisci nel carrello</span></a></p>
                                    </div>
                                </li>
                                
                            </ul>
                        </div>
                    </div>
                </div>
                <!--#include file="inc_sx_prodotti.asp"-->
            </div>
        </div>
        <!--#include file="inc_footer.asp"-->
    </body>
</html>
