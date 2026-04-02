<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top">	
<%
Response.Write "<table class='tab_paginacao max_width remover_bordas_padding'>"
Response.Write "<tr style='height: 26px'>"
Response.Write "<td class='esquerda' style='width: 33%'>"
if Request.QueryString("veio_de") = "reserva" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkReservas(parent.hiddenFrame.modo_busca);""><img src='imagens/icon-small-back.png' class='transparent-icon' alt=''>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "bibcurso" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkBibCurso(parent.hiddenFrame.modo_busca, 'detalhe');""><img src='imagens/icon-small-back.png' class='transparent-icon' alt=''>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "sels" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkSelecao(parent.hiddenFrame.modo_busca);""><img src='imagens/icon-small-back.png' class='transparent-icon' alt=''>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "favoritos" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkFavoritos(parent.hiddenFrame.modo_busca, "&Request.QueryString("listaSelecionada")&");""><span class='transparent-icon removerBordas span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "link_detalhe" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
else
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:volta_resultado(parent.hiddenFrame.modo_busca,"&Request.QueryString("pagina")&");""><img src='imagens/icon-small-back.png' class='transparent-icon' alt=''>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
end if
Response.Write "</td>"
Response.Write "<td class='centro' style='width: 34%'>"
if Request.QueryString("veio_de") = "sels" then
	%><!-- #include file="navegador_sels.asp" --><%
elseif Request.QueryString("veio_de") = "reserva" OR Request.QueryString("veio_de") = "leg" OR Request.QueryString("veio_de") = "periodico" OR Request.QueryString("veio_de") = "link_detalhe" then
	Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
else
	%><!-- #include file="navegador_detalhes.asp" --><%
end if
Response.Write "</td>"
Response.Write "<td class='direita' style='width: 33%'>"
if (somente_detalhe = false) then
	Response.Write "<a class='link_serv_custom' title='"&getTermo(global_idioma, 1345, "Clique aqui para fazer uma nova consulta.", 0)&"' href='index.asp?modo_busca="&StrQueryString("modo_busca")&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"' onClick=Resetar();><img class='transparent-icon' src='imagens/icon-small-newsearch.png' alt=''>&nbsp;"&getTermo(global_idioma, 1344, "Nova pesquisa", 0)&"</a>"
end if
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
Response.Write "</tr></table>"
Select Case cInt(Tipo)
	Case 0
		desc_tipo = getTermo(global_idioma, 1684, "Ficha completa do periódico", 0)
	
	Case 1
		desc_tipo = getTermoHtml(global_idioma, 1637, "Ficha completa da obra", 0)
		
	Case 3
		desc_tipo = getTermoHtml(global_idioma, 6286, "Ficha completa da analítica", 0)
End Select
%>
<table class="max_width remover_bordas_padding" style="border-color: #CCCCCC; table-layout: fixed;">
<colgroup>
    <col style="width: 100px" />
    <col style="width: 100px" />
    <col style="width: 100px" />
    <col style="width: auto" />
</colgroup>
<tr class="tr-abas-detalhe">
	<td class="td_center_middle td_abas_detalhe_aacr2" id="tdx_a2"><a class="link_abas" href="#" onClick="habilitaMarc(1,parent.hiddenFrame.layerX,'detalhes',<%=codigo_obra%>,<%=Tipo%>);" id="lk_a2"><%=getTermo(global_idioma, 1032, "Detalhes", 0)%></a></td>
	<td class="td_center_middle td_abas_detalhe_marc" id="tdx_c2"><a class="link_abas" href="#" onClick="habilitaMarc(1,parent.hiddenFrame.layerX,'marcTags',<%=codigo_obra%>,<%=Tipo%>);" id="lk_c2"><%=getTermo(global_idioma, 997, "MARC tags", 0)%></a></td>
	<td class="td_right_bottom">&nbsp;</td>
</tr>
<tr>
	<td class="td_right_bottom td_abas_detalhe_aacr2_esquerda" id="tdx_fm1">&nbsp;</td>
	<td class="td_right_bottom td_abas_detalhe_aacr2_centro" id="tdx_fm2">&nbsp;</td>
	<td class="td_right_bottom td_abas_detalhe_aacr2_centro" id="tdx_fm3">&nbsp;</td>
	<td class="td_right_bottom td_abas_detalhe_aacr2_direita" id="tdx_fm4">&nbsp;</td>
</tr>
<tr>
<td colspan="4" class="td_center_top ficha_marc">&nbsp;<b><%=desc_tipo%></b>&nbsp;</td>
</tr>
<tr><td colspan="4" class="td_center_top td_marc"><br />
<%
Response.Write "<div style='width: 98%; display: inline-table;' id='dfichaMarc'><img src='imagens/mozilla_blu.gif' alt=''>&nbsp;"&getTermo(global_idioma,32,"Carregando",0)&"...</div>"
%>
</td>
</tr>
</table>
</td>
</tr>
</table>
<% if Request.QueryString("ve_marc") = "div_marc" then %>
	<script type="text/javascript">
		var cod  = <%=codigo_obra%>;
		var tipo = <%=Tipo%>;
		ajxCall("ajxFichaMarc", cod, tipo, global_ficha_marc);
		global_ficha_marc = 1;
	</script>
<% end if %>