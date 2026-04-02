<% 
	sDiretorioArq="asp" 
	nao_imprime_variaveis_globais = "1"
%>

<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<%
	iIndexSrv = IntQueryString("BuscaSrv", 1)
	
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	
	sXMLFichas = ""
	
	'Variáveis utilizadas na chamada da rotina "Paginacao"
	pagina = 1
	indice = 1
	
	' Marca se a pesquisa retornou algum resultado
	bResultOk = True
	
	if len(trim(Request.QueryString("cod"))) > 0 AND Request.QueryString("veio_de") = "linkAutInfo" then
		aut_pag = Request.QueryString("cod")
	elseif Request.Form("submeteu") = "aut" AND Request.QueryString("veio_de") <> "paginacao" then
		modo_busca = "rapida"
		%><!-- #include file="monta_busca_autoridades.asp" --><%
		aut_pag = vetor_pag_auts(1)
	else
		sMsgResult = "<div class='centro'><span style='color: red'>"&getTermo(global_idioma, 1341, "Nenhum registro encontrado.", 0)&"</div>"
		bResultOk = False
	end if
%>

	<input type="hidden" id="qtde_resultados_<%=iIndexSrv%>" value="<%=GetSessao(global_cookie,"nrows_auts"&iIndexSrv)%>" />

<%
	'Cria tabela de paginação
	Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
	Response.Write "	<tr style='height: 26px'>"
	Response.Write "		<td class='esquerda' style='width: 560px'>"
	Response.Write "			&nbsp;&nbsp;"
	
	modo_busca = "rapida"
	Paginacao GetSessao(global_cookie,"nrows_auts"&iIndexSrv),global_num_linhas,global_max_links,pagina,indice,modo_busca,"autoridades",GetSessao(global_cookie,"nrows_auts"&iIndexSrv)
	
	Response.Write "		</td>"
	Response.Write "		<td class='direita'><a class='link_serv_custom' title='"&getTermo(global_idioma, 1345, "Clique aqui para fazer uma nova consulta.", 0)&"' href=""javascript:novaPesquisaAutoridade(parent.hiddenFrame.modo_busca,0);""><span class='transparent-icon span_imagem icon_16 icon-small-newsearch'></span>&nbsp;"&getTermo(global_idioma, 1344, "Nova pesquisa", 0)&"</a>"
	Response.Write "			&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "		</td>"
	Response.Write "	</tr>"
	Response.Write "	</table>"
	
	'Se a pesquisa retornar algum resultado, montar as fichas dos registros encontrados
	if(bResultOk = True)then	
		%> <!-- #include file='grid_auts.asp' --> <%
		
		Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
		Response.Write "	<tr style='height: 26px'>"
		Response.Write "		<td class='esquerda' style='width: 560px'>"
		Response.Write "			&nbsp;&nbsp;"
	
		Paginacao GetSessao(global_cookie,"nrows_auts"&iIndexSrv),global_num_linhas,global_max_links,pagina,indice,modo_busca,"autoridades",GetSessao(global_cookie,"nrows_auts"&iIndexSrv)
		
		Response.Write "		</td>"
		Response.Write "		<td class='direita'>"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1345, "Clique aqui para fazer uma nova consulta.", 0)&"' href=""javascript:novaPesquisaAutoridade(parent.hiddenFrame.modo_busca,0);""><span class='transparent-icon span_imagem icon_16 icon-small-newsearch'></span>&nbsp;"&getTermo(global_idioma, 1344, "Nova pesquisa", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write "		</td>"
		Response.Write "	</tr>"
		Response.Write "	</table>"
	else
		Response.Write sMsgResult
	end if
%>