<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top">
<%
iIndexSrv = IntQueryString("Servidor", 1)
'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%	

sPagCorrenteAbas = Request.QueryString("pagina")
if sPagCorrenteAbas = "" then	
	For iPag = 0 to Servidores.ServList.Count-1
	%>
		<script type="text/javascript">
			parent.hiddenFrame.arPagCorrente[<%=iPag%>] = 1;
			parent.hiddenFrame.arIndCorrente[<%=iPag%>] = 1;
		</script>
	<%
		sPagCorrenteAbas = "|1" + sPagCorrenteAbas
	Next		
end if

sIndCorrenteAbas = Request.QueryString("indice")
if sIndCorrenteAbas = "" then
	sIndCorrenteAbas = sPagCorrenteAbas
end if

arPaginas = Split(sPagCorrenteAbas, "|")
arIndices = Split(sIndCorrenteAbas, "|")

'Variáveis utilizadas na chamada da rotina "Paginacao"
pagina = arPaginas(iIndexSrv)
indice = arIndices(iIndexSrv)
modoBusca = GetModo_Busca

if Request.QueryString("veio_de") = "fechou" or Request.QueryString("veio_de") = "abriu" OR Request.QueryString("veio_de") = "paginacao" then

	vet_pag = right(Request.QueryString("vetor_pag"),len(Request.QueryString("vetor_pag"))-1)
	vet_pag = left(vet_pag,len(vet_pag)-1)
	a_pag = right(Request.QueryString("art_pag"),len(Request.QueryString("art_pag"))-1)
	a_pag = left(a_pag,len(a_pag)-1)

elseif (modoBusca = "legislacao") then
    %><!-- #include file='monta_busca_link_legislacao.asp' --><%
	
	if (global_versao = vSOPHIA) and (global_busca_facetada = 1) then
        sDisplay = "table-cell"
    else
        sDisplay = "block"
    end if
else
	
	%><!-- #include file='monta_busca_link.asp' --><%
	
	arPagSrv = Split(vet_pag, "|")	
	vet_pag = arPagSrv(iIndexSrv)

    if (global_versao = vSOPHIA) and (global_busca_facetada = 1) then
        sDisplay = "table-cell"
    else
        sDisplay = "block"
    end if
end if

'******************************************************************************************
arr_ordenacao = split(GetSessao(global_cookie,"ORDER_BY"),"-")
comb_ordenacao = arr_ordenacao(0)
comb_asc_desc = arr_ordenacao(1)
'******************************************************************************************
if sXMLFichas = "" then
	Set ROService = ROServer.CreateService("Web_Consulta")
    Set RetornoMontaFichas = ROService.MontaFichas(vet_pag, "", false, global_idioma)
	sXMLFichas = RetornoMontaFichas.sMsg
    Set RetornoMontaFichas = nothing
	Set ROService = nothing
end if
'******************************************************************************************

%>
<table class="max_width remover_bordas_padding">
<tr>
<td class="td_center_top" >
<%
if len(trim(sXMLFichas)) = 0 then
	if len(trim(sMsgResult)) > 0 then
		Response.Write sMsgResult
	else
		Response.Write "<span class='span_imagem icon_16 alerta0'></span>&nbsp;<span style='color: #006699'>"&getTermo(global_idioma, 1384, "Esta sessão expirou.", 0)&" "&getTermo(global_idioma, 1385, "Por favor, refaça sua busca.", 0)&"</span>"
	end if
else

    ' Busca Facetada
    if (global_versao = vSOPHIA) and (global_busca_facetada = 1) then %>

		<div id='facet_aba_<%=iIndexSrv%>' style='width: 180px; display: <%=sDisplay%>; padding-right: 0px;'>

		<table class='tab-busca-facetada max_width remover_bordas_padding centro'>
            <tr style='height: 26px'><td class="td_faceta_cabecalho">&nbsp;Filtros</td></tr>
	    	<tr><td class='td_facetas'>
		        <span class='span_imagem div_imagem_right icon_16 mozilla_blu '></span><span><%=getTermo(global_idioma, 32, "Carregando", 0)%>...</span>
            </td></tr>
		</table>

		</div>

<%  end if

    Response.Write "<div id='p_div_aba"&iIndexSrv&"' style='display:"&sDisplay&"; min-height: 288px; padding-left: 10px'>"

	Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
	Response.Write "	<tr style='height: 26px'>"
	Response.Write "		<td class='esquerda'>&nbsp;&nbsp;"
	if (global_versao = vSOPHIA) and (global_esconde_menu = 0) then 	
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1351, "Clique aqui para selecionar todos os títulos desta página.", 0)&"' href=""javascript:fichas_marcar_todos("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-check'></span>"&getTermo(global_idioma, 1348, "Selecionar todos", 0)&"</a>&nbsp;&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1352, "Clique aqui para desmarcar todos os títulos desta página.", 0)&"' href=""javascript:fichas_desmarcar_todos("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-delete'></span> "&getTermo(global_idioma, 1349, "Desmarcar selecionados", 0)&"</a>&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1353, "Clique aqui para enviar os títulos selecionados para Minha seleção.", 0)&"' href=""javascript:enviar_minha_selecao("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-selecao'></span> "&getTermo(global_idioma, 1350, "Enviar para minha seleção", 0)&"&nbsp;&nbsp;"	
		if global_hab_servicos = 1 then
			if Session("Logado") = "sim" then
				if config_multi_servbib = 1 then
					if CStr(Session("Servidor_Logado")) = CStr(iIndexSrv) then
						logado = "true"
                        sEstadoIconeFavorito = "transparent-icon icon-small-star"
					else
						logado = "false"
                        sEstadoIconeFavorito = "icon-small-star-x"
					end if
				else
					logado = "true"
                    sEstadoIconeFavorito = "transparent-icon icon-small-star"
				end if
			else
				logado = "false"
                sEstadoIconeFavorito = "icon-small-star-x"
			end if	
			Response.Write "<a class='link_serv_custom' title='"&getTermo(global_idioma, 8318, "Clique aqui para salvar os títulos selecionados nos favoritos.", 0)&"' href=""javascript:salvar_favoritos("&iIndexSrv&",'resultado',"&logado&");"">"
			Response.Write "<span class='span_imagem div_imagem_right_3 icon_16 "&sEstadoIconeFavorito&"'></span>"&getTermo(global_idioma, 8317, "Salvar favoritos", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"	
		end if
	end if
	Response.Write "		</td>"

	Response.Write "		</td>"
	Response.Write "		<td class='direita' style='width: 140px'>"
	Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1345, "Clique aqui para fazer uma nova consulta.", 0)&"' href=""javascript:novaPesquisa(parent.hiddenFrame.modo_busca);"">"
	Response.Write "			<span class='transparent-icon span_imagem icon_16 icon-small-newsearch'></span>&nbsp;"&getTermo(global_idioma, 1344, "Nova pesquisa", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "		</td>"
		
	Response.Write "	</tr>"
	Response.Write "	</table>"

    Response.Write "<div id='p_div_aba"&iIndexSrv&"_resultado' style='min-height: 261px'>"
	
	'Cria tabela de paginação
	Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
	Response.Write "	<tr style='height: 26px'>"
	Response.Write "		<td class='esquerda'>&nbsp;&nbsp;"
	
	Paginacao GetSessao(global_cookie,"nrows"&iIndexSrv),global_num_linhas,global_max_links,pagina,indice,modo_busca,"resultado",GetSessao(global_cookie,"nrows_real"&iIndexSrv)
	
	Response.Write "		</td>"
	Response.Write "	</tr>"
	Response.Write "	</table>"

	sOrigem = "Pesquisa"
	%> <!-- #include file='grid_ficha.asp' --> <%
	
	Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
	Response.Write "	<tr style='height: 26px'>"
	Response.Write "		<td class='esquerda'>&nbsp;&nbsp;"
	
	Paginacao GetSessao(global_cookie,"nrows"&iIndexSrv),global_num_linhas,global_max_links,pagina,indice,modo_busca,"resultado",GetSessao(global_cookie,"nrows_real"&iIndexSrv)
	
	Response.Write "		</td>"
	Response.Write "	</tr>"
	Response.Write "	</table>"		

	'Cria menu de ações (Selecionar, Marcar Todos, etc)
	Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
	Response.Write "	<tr style='height: 26px'>"
	Response.Write "		<td class='esquerda'>&nbsp;&nbsp;"
	if (global_versao = vSOPHIA) and (global_esconde_menu = 0) then
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1351, "Clique aqui para selecionar todos os títulos desta página.", 0)&"' href=""javascript:fichas_marcar_todos("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-check'></span>"&getTermo(global_idioma, 1348, "Selecionar todos", 0)&"</a>&nbsp;&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1352, "Clique aqui para desmarcar todos os títulos desta página.", 0)&"' href=""javascript:fichas_desmarcar_todos("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-delete'></span> "&getTermo(global_idioma, 1349, "Desmarcar selecionados", 0)&"</a>&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1353, "Clique aqui para enviar os títulos selecionados para Minha seleção.", 0)&"' href=""javascript:enviar_minha_selecao("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-selecao'></span> "&getTermo(global_idioma, 1350, "Enviar para minha seleção", 0)&"&nbsp;&nbsp;"	
		if global_hab_servicos = 1 then
			if Session("Logado") = "sim" then
				if config_multi_servbib = 1 then
					if CStr(Session("Servidor_Logado")) = CStr(iIndexSrv) then
						logado = "true"
                        sEstadoIconeFavorito = "transparent-icon icon-small-star"
					else
						logado = "false"
                        sEstadoIconeFavorito = "icon-small-star-x"
					end if
				else
					logado = "true"
                    sEstadoIconeFavorito = "transparent-icon icon-small-star"
				end if
			else
				logado = "false"
                sEstadoIconeFavorito = "icon-small-star-x"
			end if
			Response.Write "<a class='link_serv_custom' title='"&getTermo(global_idioma, 8318, "Clique aqui para salvar os títulos selecionados nos favoritos.", 0)&"' href=""javascript:salvar_favoritos("&iIndexSrv&",'resultado',"&logado&");"">"
			Response.Write "<span class='span_imagem div_imagem_right_3 icon_16 "&sEstadoIconeFavorito&"'></span>"&getTermo(global_idioma, 8317, "Salvar favoritos", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"	
		end if
	end if
	Response.Write "		</td>"		
	
	Response.Write "		<td class='direita' style='width: 140px'><a class='link_serv_custom' title='"&getTermo(global_idioma, 1345, "Clique aqui para fazer uma nova consulta.", 0)&"' href=""javascript:novaPesquisa(parent.hiddenFrame.modo_busca);"">"
	Response.Write "		<span class='transparent-icon span_imagem icon_16 icon-small-newsearch'></span>&nbsp;"&getTermo(global_idioma, 1344, "Nova pesquisa", 0)&"</a>"
	Response.Write "			&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "		</td>"
	Response.Write "	</tr>"
	Response.Write "	</table>"
	'---------------------------------------

    Response.Write "</div>"
    Response.Write "</div>"

    ' Busca Facetada
    if (global_versao = vSOPHIA) and (global_busca_facetada = 1) then
    %>
        <script type="text/javascript">
            $(document).ready(function() {
                <% if (modoBusca = "rapida") then %>
                    MontarBuscaFaceta(1, <%=iIndexSrv%>, '');
                <% elseif (modoBusca = "combinada") then %>
                    MontarBuscaFaceta(3, <%=iIndexSrv%>, '');
                <% elseif (modoBusca = "legislacao") then %>
                    MontarBuscaFaceta(5, <%=iIndexSrv%>, '');
                <% end if %>
            });
		</script>
    <%
    end if
end if
%>
</tr>
</table>
</td>
</tr>
</table>
