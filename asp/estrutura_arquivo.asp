<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top">	
<%
Response.Write "<table class='tab_paginacao max_width remover_bordas_padding'>"
Response.Write "<tr style='height: 26px'>"
Response.Write "<td class='esquerda' style='width: 33%'>"
if Request.QueryString("veio_de") = "reserva" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkReservas(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "bibcurso" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkBibCurso(parent.hiddenFrame.modo_busca,'detalhe');""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "sels" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkSelecao(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "favoritos" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkFavoritos(parent.hiddenFrame.modo_busca, "&Request.QueryString("listaSelecionada")&");""><span class='transparent-icon removerBordas span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "link_detalhe" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
else
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:volta_resultado(parent.hiddenFrame.modo_busca,"&Request.QueryString("pagina")&");""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
end if
Response.Write "</td>"
Response.Write "<td class='centro' style='width: 34%'>"
if Request.QueryString("veio_de") = "sels" then
	%><!-- #include file="navegador_sels.asp" --><%
elseif Request.QueryString("veio_de") = "reserva" OR Request.QueryString("veio_de") = "leg" OR Request.QueryString("veio_de") = "periodico" OR Request.QueryString("veio_de") = "link_detalhe" OR Request.QueryString("veio_de") = "favoritos" then
	Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
else
	%><!-- #include file="navegador_detalhes.asp" --><%
end if
Response.Write "</td>"
Response.Write "<td class='direita' style='width: 33%'>"
if (somente_detalhe = false) then
	Response.Write "<a class='link_serv_custom' title='"&getTermo(global_idioma, 1345, "Clique aqui para fazer uma nova consulta.", 0)&"' href='index.asp?modo_busca="& GetModo_Busca &"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"' onClick=Resetar();><span class='transparent-icon span_imagem icon_16 icon-small-newsearch'></span>&nbsp;"&getTermo(global_idioma, 1344, "Nova pesquisa", 0)&"</a>"
end if
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
Response.Write "</tr></table>"
%>
<table class="remover_bordas_padding" style="table-layout: fixed; border-color: #CCCCCC; width: 100%">
<colgroup>
    <col style="width: 100px" />
    <col style="width: 100px" />
	<col style="width: 100px" />
    <% if (possui_estrutura_hierarquica) then %>
		<col style="width: 100px" />
	<% end if %>
    <col style="width: auto" />
</colgroup>
<tr class="tr-abas-detalhe">
    <%
        'BND
        if (global_numero_serie = 5516) then
            FichaResumida = "Detalhes"
            ocultaItem = "display: none;"
            exibeBorda = "border-top: 1px solid #999 !important;"
            ocultaBorda = "border-top: 0 !important;"
        else
            FichaResumida = getTermo(global_idioma, 1032, "Detalhes", 0)
            exibeBorda = ""
            ocultaBorda = ""
        end if
    %>
	<td class="td_center_middle td_abas_detalhe_aacr2" id="tdx_a4"><a class="link_abas" href="#" onClick="habilitaMarc(1,parent.hiddenFrame.layerX,'detalhes',<%=codigo_obra%>,<%=Tipo%>);" id="lk_a4"><%=FichaResumida%></a></td>
	<td class="td_center_middle td_abas_detalhe_marc" id="tdx_b4"><a class="link_abas" href="#" onClick="habilitaMarc(1,parent.hiddenFrame.layerX,'marcTags',<%=codigo_obra%>,<%=Tipo%>);" id="lk_b4"><%=getTermo(global_idioma, 997, "MARC tags", 0)%></a></td>
	<td class="td_center_middle td_abas_detalhe_marc" id="tdx_dc4"><a class="link_abas" href="#" onclick="habilitaMarc(1,parent.hiddenFrame.layerX,'dublinCore',<%=codigo_obra%>,<%=Tipo%>);" id="lk_dc4">Dublin Core</a></td>
	<td class="td_center_middle td_abas_detalhe_marc" id="tdx_c4"><a class="link_abas" href="#" onclick="habilitaMarc(0,parent.hiddenFrame.layerX,'arquivo',<%=codigo_obra%>,<%=Tipo%>);" id="lk_c4"><%=getTermo(global_idioma, 7482, "Registros relacionados", 0)%></a></td>
	<td class="td_right_bottom">&nbsp;</td>
</tr>
<tr>
	<td class="td_right_bottom td_abas_detalhe_aacr2_esquerda" id="tdx_ar1">&nbsp;</td>
	<td class="td_right_bottom td_abas_detalhe_aacr2_centro" id="tdx_ar2">&nbsp;</td>
	<td class="td_right_bottom td_abas_detalhe_aacr2_centro" id="tdx_ar5">&nbsp;</td>
	<td class="td_right_bottom td_abas_detalhe_aacr2_centro" id="tdx_ar3">&nbsp;</td>
	<td class="td_right_bottom td_abas_detalhe_aacr2_direita" id="tdx_ar4">&nbsp;</td>
</tr>
<tr>
<% if (possui_estrutura_hierarquica) then %>
	<td colspan="5" class='td_center_top estrutura_arquivo'>&nbsp;<b><%=getTermo(global_idioma, 7482, "Registros relacionados", 0)%></b>&nbsp;</td>
	</tr>
	<tr><td colspan="5" class="td_center_top td_marc" style="background-color: #fff;"><br />
<% else %>
	<td colspan="4" class='td_center_top estrutura_arquivo'>&nbsp;<b><%=getTermo(global_idioma, 7482, "Registros relacionados", 0)%></b>&nbsp;</td>
	</tr>
	<tr><td colspan="4" class="td_center_top td_marc" style="background-color: #fff;"><br />
<% end if %>
<%
Response.Write "<div id='divArquivo'><span class='span_imagem div_imagem_right icon_16 mozilla_blu '></span>&nbsp;"&getTermo(global_idioma,32,"Carregando",0)&"...</div>"
%>
</td>
</tr>
</table>
</td>
</tr>
</table>
<% if Request.QueryString("ve_marc") = "div_arquivo" then %>	
	<script type="text/javascript">
		var cod = <%=codigo_obra%>;
		var tipo = <%=Tipo%>;
		ajxCall("ajxArquivo", cod, tipo, global_arquivo_load);
		global_arquivo_load = 1;
	</script>
<% end if %>