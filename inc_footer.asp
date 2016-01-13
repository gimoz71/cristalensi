<footer id="footer">
    <p class="credits">[<a href="http://www.viadeimedici.it" target="_blank" title="realizzazione siti internet e web marketing">realizzazione sito internet e web marketing</a>] - [<a href="http://www.matteopalpacelli.com" target="_blank" title="web marketing e seo a Firenze">seo</a>]</p>
</footer>

<script src="/js/init.js"></script>

<!-- Load the script -->
<script src="/js/cookies-enabler.js"></script>

<!-- Init the script -->
<script>
	COOKIES_ENABLER.init({
		scriptClass: 'ce-script',
		iframeClass: 'ce-iframe',

		acceptClass: 'ce-accept',
		dismissClass: 'ce-dismiss',
		disableClass: 'ce-disable',

		bannerClass: 'ce-banner',
		bannerHTML:
			'<div style="text-align:center">Questo sito internet utilizza i cookies per migliorare la vostra esperienza online. '
				+'<a href="privacy.asp" class="info">'
					+'Informativa estesa'
				+'</a>'
				+'<a href="#" class="ce-accept">'
					+'Accetta le condizioni'
				+'</a>'
			+'</div>',

		eventScroll: true,
		scrollOffset: 200,

		clickOutside: true,

		cookieName: 'ce-cookie',
		cookieDuration: '365',

		iframesPlaceholder: true,
		iframesPlaceholderHTML:
			'<p>Per visualizzare questo contenuto devi'
				+'<a href="#" class="ce-accept">accettare i Cookies</a>'
			+'</p>',
		iframesPlaceholderClass: 'ce-iframe-placeholder',

		// Callbacks
		onEnable: '',
		onDismiss: '',
		onDisable: ''
	});
</script>
