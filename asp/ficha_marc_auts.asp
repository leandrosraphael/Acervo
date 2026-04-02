<%
pagina = Request.QueryString("pagina")
codigo_aut = Request.QueryString("codigo_aut")
tipo_autoridade = request.QueryString("tipo_autoridade")
'Define o Servbib padrão da obra
iIndexSrv = IntQueryString("Servidor", 1)
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

%>
<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top">	
<%
Response.Write "<table class='tab_paginacao max_width remover_bordas_padding'>"
Response.Write "<tr style='height: 26px'>"
Response.Write "<td class='esquerda' style='width: 33%'>"
if Request.QueryString("veio_de") = "linkAutInfo" then
	funcao_js = "volta_autoridade('linkAutInfo',"&Request.QueryString("pagina")&");"
elseif Request.QueryString("veio_de") = "detalhes" then
	funcao_js = "javascript:history.go(-1)"	
else
	funcao_js = "volta_autoridade(parent.hiddenFrame.modo_busca,"&Request.QueryString("pagina")&");"
end if
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:"&funcao_js&";""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
Response.Write "</td>"
Response.Write "<td class='td_center_middle' style='width: 34%'>"
if Request.QueryString("veio_de") <> "linkAutInfo" AND Request.QueryString("veio_de") <> "detalhes" then
	%><!-- #include file="navegador_auts.asp" --><%
else
	Response.Write "&nbsp;"
end if
Response.Write "</td>"
Response.Write "<td class='direita' style='width: 33%'>"
Response.Write "<a class='link_serv_custom' title='"&getTermo(global_idioma, 1345, "Clique aqui para fazer uma nova consulta.", 0)&"' href=""javascript:LinkAutoridades(parent.hiddenFrame.modo_busca,0);""><span class='transparent-icon span_imagem icon_16 icon-small-newsearch'></span>&nbsp;"&getTermo(global_idioma, 1344, "Nova pesquisa", 0)&"</a>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
Response.Write "</tr></table>"
%>
<table class="remover_bordas_padding" style="table-layout: fixed; width: 100%; border-color: #CCCCCC">
<colgroup>
    <col style="width: 100px;" />
    <col style="width: 100px;" />
    <col style="width: 100px;" />
    <col style="width: auto;" />
</colgroup>
<tr class="tr-abas-detalhe" id="detalhe-autoridade-aacr2">
	<td class="td_bottom_center td_abas_detalhe_aacr2" id="div-aacr_aacr2"><a class="link_abas" href="#" onClick="TrocarAbaDetalheAutoridade('aacr2')" id="link-aacr_aacr2"><%=getTermo(global_idioma, 1557, "Ficha completa", 0)%></a></td>
    <%if (global_marc = 1) then %>
	   <td class="td_bottom_center td_abas_detalhe_marc" id="div-aacr_marc"><a class="link_abas" href="#" onClick="TrocarAbaDetalheAutoridade('marc')" id="link-aacr_marc"><%=getTermo(global_idioma, 997, "MARC tags", 0)%></a></td>
        <%if (global_recursos_avancados = 1) and (tipo_autoridade = "150") then%>
	         <td class="td_bottom_center td_abas_detalhe_tesauro" id="div-aacr_tesauro"><a class="link_abas" href="#" onClick="TrocarAbaDetalheAutoridade('tesauro')" id="link-aacr_tesauro"><%=getTermo(global_idioma, 470, "Tesauro", 0)%></a></td>
        <% else %>
            <td class="td_right_bottom">&nbsp;</td>
        <% end if %>
    <%else %>
        <%if (global_recursos_avancados = 1) and (tipo_autoridade = "150") then%>
	         <td class="td_bottom_center td_abas_detalhe_tesauro" id="div-aacr_tesauro"><a class="link_abas" href="#" onClick="TrocarAbaDetalheAutoridade('tesauro')" id="link-aacr_tesauro"><%=getTermo(global_idioma, 470, "Tesauro", 0)%></a></td>
        <% else %>
            <td class="td_right_bottom">&nbsp;</td>
        <% end if %>
        <td class="td_right_bottom">&nbsp;</td>
    <%end if %>
	<td class="td_right_bottom">&nbsp;</td>
</tr>
<tr>
	<td colspan="4" class="td_right_bottom td_abas_detalhe_aacr2_esquerda">&nbsp;</td>
</tr>
<tr>
<%

if isnumeric(codigo_obra) then
	'--------------------------------------------------------------------------
	'                             INICIO DA FICHA MARC
	'--------------------------------------------------------------------------
	
	'Destaque dos termos pesquisados
	redim array_campos (5)
	redim array_palavras (5)
	redim array_frase_exata (5)
	Set objParamDestaca = ROServer.CreateComplexType("TParamBuscaHighlight")
	
	array_campos(0) = GetPosCampoPesquisa(Request.QueryString("campo1"))
	array_campos(1) = "0"
	array_campos(2) = "0"
	array_campos(3) = "0"
	array_campos(4) = "0"
		
	sValor = RemoveUnderline(Request.QueryString("valor1"))
	array_palavras(0) = SemAspas(sValor)
	array_frase_exata(0) = EntreAspas(sValor)
		
	array_palavras(1) = ""
	array_frase_exata(1) = ""
		
	array_palavras(2) = ""
	array_frase_exata(2) = ""
		
	array_palavras(3) = ""
	array_frase_exata(3) = ""
		
	array_palavras(4) = ""
	array_frase_exata(4) = ""
				
	'CRIANDO ARRAYS RO
	Set aiCampos = ROServer.CreateComplexType("TInteiro")
	aiCampos.Param0 = array_campos(0)
	aiCampos.Param1 = array_campos(1)
	aiCampos.Param2 = array_campos(2)
	aiCampos.Param3 = array_campos(3)
	aiCampos.Param4 = array_campos(4)
		
	Set aiPalavras = ROServer.CreateComplexType("TString")
	aiPalavras.Param0 = array_palavras(0)
	aiPalavras.Param1 = array_palavras(1)
	aiPalavras.Param2 = array_palavras(2)
	aiPalavras.Param3 = array_palavras(3)
	aiPalavras.Param4 = array_palavras(4)
	
	Set aiFrase = ROServer.CreateComplexType("TString")
	aiFrase.Param0 = array_frase_exata(0)
	aiFrase.Param1 = array_frase_exata(1)
	aiFrase.Param2 = array_frase_exata(2)
	aiFrase.Param3 = array_frase_exata(3)
	aiFrase.Param4 = array_frase_exata(4)

	objParamDestaca.asPalavras   = aiPalavras
	objParamDestaca.asFraseExata = aiFrase
	objParamDestaca.aiCampos     = aiCampos
	
	'Fim do destaque os termos pesquisados	
	
	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	xml_ficha_marc = ROService.MontaAutoridadesMarc(codigo_aut,global_idioma, objParamDestaca)
	SET ROService = nothing
	TrataErros(1)


	ficha = Formata_Ficha(xml_ficha_marc,"auts",global_idioma,iIndexSrv)

	if len(trim(xml_ficha_marc)) < 4  OR left(trim(uCase(xml_ficha_marc)),4) = "ERRO" then
		Response.write "<td colspan='4' class='td_center_top ficha_detalhes'>&nbsp;<b>" & getTermo(global_idioma, 1686, "Ficha da autoridade", 0) & "</b>&nbsp;</td>"
		Response.write "</tr>"
		Response.write "<tr><td colspan='4' class='td_center_top td_marc'><br />"
		Response.write "<table style='width:98%; display: inline-table;' class='remover_bordas_padding'>"
		Response.write "<tr>"
		Response.Write "<tr><td><em>"&getTermo(global_idioma, 1329, "Desculpe, houve um erro na comunicação com o servidor.", 0)&"<br />"&getTermo(global_idioma, 1333, "Por favor, tente novamente mais tarde.", 0)&"</em><br /> "&xml_ficha&"</td></tr>"
		Response.End()
	else
		tipoAut = ""
		descAut = ""
		posTipoIni = InStr(ficha,"<!--tipo:")

		If (posTipoIni > 0) then
			posTipoIni = posTipoIni + 9
			posTipoFim = InStr(ficha,"/tipo-->")
			tipoAut = Mid(ficha, posTipoIni, posTipoFim - posTipoIni) 

			select case tipoAut
				case "100"
					descAut = getTermo(global_idioma, 732, "Pessoa", 0)
				case "110"
					descAut = getTermo(global_idioma, 62, "Instituição", 0)
				case "111"
					descAut = getTermo(global_idioma, 63, "Evento", 0)
				case "130"
					descAut = getTermo(global_idioma, 70, "Título uniforme", 0)
				case "148"
					descAut = getTermo(global_idioma, 71, "Termo cronológico", 0)
				case "150"
					descAut = getTermo(global_idioma, 742, "Termo tópico", 0)
				case "151"
					descAut = getTermo(global_idioma, 1148, "Local geográfico", 0)
				case "155"
					descAut = getTermo(global_idioma, 3707, "Termo de gênero e forma", 0)
				case "180"
					descAut = getTermo(global_idioma, 75, "Subdivisão geral", 0)
				case "181"
					descAut = getTermo(global_idioma, 76, "Subdivisão geográfica", 0)
				case "182"
					descAut = getTermo(global_idioma, 77, "Subdivisão cronológica", 0)
				case "185"
					descAut = getTermo(global_idioma, 3390, "Subdivisão de forma", 0)
				case else
					descAut = ""
			End Select

		end if
		

		naoRevisado = InStr(ficha,"<!--nao_revisado-->") > 0

		Response.write "<td colspan='4' class='td_center_top ficha_detalhes'>&nbsp;<b>" & getTermo(global_idioma, 1686, "Ficha da autoridade", 0) & " - "&descAut& "</b>&nbsp;</td>"
		Response.write "</tr>"
		Response.write "<tr><td colspan='4' class='td_center_top td_marc'><br />"
		if naoRevisado then
			Response.write "<div style='text-align:right; color:#f00; margin-right: 10px; margin-bottom: 2px;'>" & getTermo(global_idioma, 7483, "Termo não autorizado", 0) & "</div>"
		end if
		Response.write "<table style='width:98%; display: inline-table;' class='remover_bordas_padding'>"
		Response.write "<tr>"
		Response.Write "<tr><td>"& ficha &"</td></tr>"
	end if
else
	Response.write "<td colspan='4' class='td_center_top ficha_detalhes'>&nbsp;<b>" & getTermo(global_idioma, 1686, "Ficha da autoridade", 0) & "</b>&nbsp;</td>"
	Response.write "</tr>"
	Response.write "<tr><td colspan='4' class='td_center_top td_marc'><br />"
	Response.write "<table style='width:98%; display: inline-table;' class='remover_bordas_padding'>"
	Response.write "<tr>"
	Response.Write "<td>ERRO</td>"
end if
response.write "</tr>" 
response.write "</table><br /><br />"
%>
</td>
</tr>
</table>
</td>
</tr>
</table>
