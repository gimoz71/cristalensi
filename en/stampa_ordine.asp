<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="strConn.asp"-->
<%
	IdOrdine=request("IdOrdine")
	if IdOrdine="" then IdOrdine=0
	
	mode=request("mode")
	if mode="" then mode=0
		
	
	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where pkid="&idOrdine
	ss.Open sql, conn, 3, 3
	
	if ss.recordcount>0 then
		TotaleCarrello=ss("TotaleCarrello")
		CostoSpedizioneTotale=ss("CostoSpedizione")
		if CostoSpedizioneTotale="" or isnull(CostoSpedizioneTotale)  then CostoSpedizioneTotale=0
		TipoTrasporto=ss("TipoTrasporto")
		DatiSpedizione=ss("DatiSpedizione")
		NoteCliente=ss("NoteCliente")
		
		FkPagamento=ss("FkPagamento")
		TipoPagamento=ss("TipoPagamento")
		CostoPagamento=ss("CostoPagamento")
		
		Nominativo=ss("Nominativo")
		Rag_Soc=ss("Rag_Soc")
		Cod_Fisc=ss("Cod_Fisc")
		PartitaIVA=ss("PartitaIVA")
		Indirizzo=ss("Indirizzo")
		Citta=ss("Citta")
		Provincia=ss("Provincia")
		CAP=ss("CAP")
		
		TotaleGenerale=ss("TotaleGenerale")
		
		DataAggiornamento=ss("DataAggiornamento")
		
	end if
	
	ss.close
	
%>
<html>
<head>
<title>CRISTALENSI - Ordine</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="stile_stampa.css" rel="stylesheet" type="text/css">
<!--Codice Statistiche Google Analytics Iury Mazzoni ## NON CANCELLARE!! ## -->
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-320952-2']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
<!--Codice Statistiche Google Analytics Iury Mazzoni ## NON CANCELLARE!! ## -->
</head>

<body <%if mode=1 then%>onLoad="print();"<%end if%> style="background-color:#FFFFFF">
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF"><tr><td align="left">
<table width="750" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF" style="border-color:#000000; border-width:1px; border-style:solid;">
					  <tr>
						<td colspan="2">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="intestazione-elenco">
							<tr>
							  <td align="left" valign="middle">
							  <h2>Cristalensi Snc<br>
							    Di Lensi Massimiliano & C.<br>
							  P.I. 0530582048<br>50056 Montelupo F.no (FI)<br>Via arti e mestieri, 1
                              <br>
							  <br>
							  </h2>							  </td>
							</tr>
						</table>					</td>
					  </tr>
					  
					  <tr>
						<td colspan="2">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="intestazione-elenco">
							<tr>
							  <td align="left" valign="middle">
							  <h2>Ordine n° <%=idordine%></h2>
							  </td>
							  <td align="right" valign="middle">Data <%=Left(DataAggiornamento, 10)%></td>
							</tr>
						</table>					</td>
					  </tr>
					  <tr>
						<td colspan="2" height="30">&nbsp;</td>
					  </tr>
					  <tr>
					  	<td colspan="2">
						
		
		
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td align="center" valign="top">

  		
			  <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                
				<tr class="sfondo-giallo">
                  <td height="20" align="left">&nbsp;[codice articolo] nome prodotto</td>
				  <td height="20" align="right">quantità</td>
				  <td height="20" align="right">costo unitario</td>
				  <td width="131" height="20" align="right">totale</td>
			    </tr>
				<tr>
                  <td colspan="4" align="left" class="intestazione-elenco"><img src="immagini/spacer.gif" height="5"></td>
                </tr>
				<tr>
                <td colspan="4" align="left">&nbsp;</td>
                </tr>
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RigheOrdine.PkId, RigheOrdine.FkOrdine, RigheOrdine.PrezzoProdotto as PrezzoProdotto, RigheOrdine.FkProdotto, RigheOrdine.Quantita, RigheOrdine.TotaleRiga, Prodotti.Titolo, Prodotti.CodiceArticolo FROM Prodotti INNER JOIN RigheOrdine ON Prodotti.PkId = RigheOrdine.FkProdotto WHERE (((RigheOrdine.FkOrdine)="&idOrdine&"))"
	rs.Open sql, conn, 1, 1
	num_prodotti_carrello=rs.recordcount
if rs.recordcount>0 then
%>
<%
Do while not rs.EOF
%>					
				  <tr>
                  <td align="left" width="341">
				  <b>[<%=rs("codicearticolo")%>]&nbsp;<%=rs("titolo")%></b>				  				  </td>
                  <td align="right" width="89">
				  <%
				  quantita=rs("quantita")
				  if quantita="" then quantita=1
				  %>
				  <%=quantita%> pezzi </td>
                  <td align="right" width="119"><%=FormatNumber(rs("PrezzoProdotto"),2)%>€</td>
                  <td align="right"><%=FormatNumber(rs("TotaleRiga"),2)%>€</td>
                  </tr>
                  <tr>
                  <td colspan="4" align="left" class="divisione-elenco"><img src="immagini/spacer.gif" height="10"></td>
                  </tr>
				  <tr>
                  <td colspan="4" align="left"><img src="immagini/spacer.gif" height="10"></td>
                  </tr>
<%
conta=conta+1
rs.movenext
loop
%>				
<%end if%>
<%rs.close%>
              </table>
			  <br>
			  <table width="100%"  border="0" cellpadding="0" cellspacing="0">
				<tr class="sfondo-giallo">
				<td width="609" height="20" align="left">&nbsp;Modalit&agrave; di spedizione </td>
				<td width="71" height="20" align="right">totale</td>
				</tr>
				<tr>
				<td colspan="2" align="left" class="intestazione-elenco"><img src="immagini/spacer.gif" height="5"></td>
				</tr>
				<tr>
				<td colspan="2" align="left">&nbsp;</td>
				</tr>
				<tr>
			    <td align="left" valign="top">
			    <b><%=TipoTrasporto%></b>				</td>
			    <td align="right" valign="top"><%=FormatNumber(CostoSpedizioneTotale,2)%>€</td>
			    </tr>
				<tr>
                <td colspan="2" align="left" class="divisione-elenco"><img src="immagini/spacer.gif" height="10"></td>
                </tr>
				<tr>
                <td colspan="2" align="left"><img src="immagini/spacer.gif" height="10"></td>
                </tr>
				<tr>
                <td height="35" colspan="2" align="left"><b>Riferimenti per l'indirizzo di spedizione:</b><br><%=DatiSpedizione%></td>
                </tr>
				<tr>
                <td colspan="2" align="left"><img src="immagini/spacer.gif" height="10"></td>
                </tr>
				<tr>
                <td height="35" colspan="2" align="left"><b>Eventuali annotazioni:</b><br><%=NoteCliente%></td>
                </tr>
				<tr>
                <td colspan="2" align="left"><img src="immagini/spacer.gif" height="10"></td>
                </tr>
			  </table>
			  <table width="100%"  border="0" cellpadding="0" cellspacing="0">
				<tr class="sfondo-giallo">
				<td width="89%" height="20" align="left">&nbsp;Modalit&agrave; di Pagamento </td>
				<td width="11%" height="20" align="right">totale</td>
				</tr>
				<tr>
				<td colspan="2" align="left" class="intestazione-elenco"><img src="immagini/spacer.gif" height="5"></td>
				</tr>
				<tr>
				<td colspan="2" align="left">&nbsp;</td>
				</tr>
				<tr>
			    <td align="left" valign="top">
			    <b><%=TipoPagamento%></b>				  				  </td>
			    <td align="right" valign="top"><%=FormatNumber(CostoPagamento,2)%>€</td>
			    </tr>
				<tr>
                <td colspan="2" align="left" class="divisione-elenco"><img src="immagini/spacer.gif" height="10"></td>
                </tr>
				<tr>
                <td height="25" colspan="2" align="left"><b>Riferimenti per i dati di fatturazione: </b></td>
                </tr>
				<tr>
                <td colspan="2" align="center">
				<table cellpadding="0" cellspacing="0" border="0" width="95%">
                    <tr align="left">
                      <td height="20" colspan="2"><%if Rag_Soc<>"" then%><%=Rag_Soc%>&nbsp;&nbsp;<%end if%><%if nominativo<>"" then%><%=nominativo%><%end if%></td>
                    </tr>
                    <tr align="left">
                      <td height="20" colspan="2">Codice fiscale: <%=Cod_Fisc%><%if PartitaIVA<>"" then%> - Partita IVA: <%=PartitaIVA%><%end if%></td>
				    </tr>
                    <tr align="left">
                      <td height="20" colspan="2"><%=indirizzo%></td>
                    </tr>
                    <tr align="left">
                      <td height="20" colspan="2"><%=cap%> - <%=citta%> (<%=provincia%>)</td>
                    </tr>
				</table>				</td>
                </tr>
				<tr>
                <td colspan="2" align="left"><img src="immagini/spacer.gif" height="10"></td>
                </tr>
			  </table>
			  <br>
  
			  <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="prodotto">
                
				 
			    <tr>
                  <td colspan="2"  align="left" class="divisione-elenco"><img src="immagini/spacer.gif" height="1"></td>
                </tr> 
				<tr>
                  <td colspan="2" align="left"><img src="immagini/spacer.gif" height="5"></td>
                </tr>
				<tr>
                  <td height="25" colspan="2" align="right" class="sfondo-giallo"><strong> Totale ordine:&nbsp;
                      <%if TotaleGenerale<>0 then%>
                      <%=FormatNumber(TotaleGenerale,2)%>
                      <%else%>
                      0,00
                      <%end if%>
                      €&nbsp;</strong></td>
                </tr>
			    <tr>
                  <td colspan="2"  align="left" class="divisione-elenco"><img src="immagini/spacer.gif" height="5"></td>
                </tr>
				<tr>
                  <td colspan="2" align="left"><img src="immagini/spacer.gif" height="10"></td>
                </tr>
				<tr>
                  <td align="left">&nbsp;</td>
				  <td align="right"><input type="button" name="stampa" value="Stampa ordine" class="button" onClick="print();">
				  </td>
				</tr>
              </table>
			  <br>
			  </td>
          </tr>
          </table>
						
						
						</td>
					  </tr>

</table>
</td></tr></table>
</body>
</html>
<!--#include file="strClose.asp"-->
