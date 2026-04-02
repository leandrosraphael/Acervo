<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top">
<table class="max_width remover_bordas_padding">
<tr>
<td class="td_resultados td_center_top">
<%
'Pega o índice do servidor selecionado no combo
iComboIndex = Trim(Request.QueryString("aut_bib"))
bMudaServidor = false

'Se (iComboIndex = ""), então está vindo da página de detalhes
if (iComboIndex = "") and (config_multi_servbib = 0) then
	'A seleção no combo de bibliotecas vem por url
	iComboIndex = Request.QueryString("iSrvCombo")
	
	'Pega o servidor corrente para ativar a aba correta
	iSrvCorrente = IntQueryString("Servidor",1)
else
	bMudaServidor = true
	if config_multi_servbib = 1 then
		if (iComboIndex = "") then
			iSrvCorrente = IntQueryString("Servidor",1)
			iComboIndex = "0"
		elseif (iComboIndex = "0") then
			iSrvCorrente = 1
		else
			if (InStr(iComboIndex, ".") > 0) then
				arServBib = Split(iComboIndex, ".")
				iSrvCorrente = arServBib(0)
			else
				iSrvCorrente = iComboIndex
			end if
		end if
	else
		iSrvCorrente = 1
	end if
%>
	<script type="text/javascript"> 
		// Marca o servidor de aplicação corrente
		parent.hiddenFrame.iSrvCorrente_Aut = <%=iSrvCorrente%>; 	
	</script>
<%	
end if

'Inicializa a variável que define a quantidade de abas (nº de servidores)
iNumAba = 1
bCriouAbas = false

'******************************************************************************************
'Início da montagem das abas 
if (config_multi_servbib = 1) and (iComboIndex = "0") then	

	Response.Write "<table class='tab_servidores max_width remover_bordas_padding'>"
	Response.Write "<tr>"
	Response.Write "  <td colspan='3'>"

	Response.Write "	<table class='max_width remover_bordas_padding direita'>" 
	Response.Write "	<tr>"
	
	iNumAba = Servidores.ServList.Count
		
	'Cria uma aba para cada servidor
	For i = 1 to iNumAba
	
		'Define o estilo da aba
		sClass = "td_abas td_abas_servidor "

		if i = 1 then
			sClass = sClass & "td_abas_esquerda "
		elseif (i = 2) or (i < iNumAba) then
			sClass = sClass & "td_abas_centro "
		else
			sClass = sClass & "td_abas_direita "
		end if

		if i = cInt(iSrvCorrente) then
			sClass = sClass & "background_aba_ativa"
		else	
			sClass = sClass & "background_aba_inativa"
		end if
		
		Response.Write "		<td class='td_center_middle "&sClass&"' id='a_td_aba"&i&"'>"
		Response.Write "			<a class='link_abas_srv' id='a_lk_aba"&i&"' href=""javascript:Mostra_Aba("&iNumAba&","&i&",'Autoridades','a');"">"

		if (Request.Form("submeteu") <> "aut") then
			Response.Write Servidores.ServList.Item(i-1).Nome
			if (GetSessao(global_cookie,"nrows_auts"&i) <> "") then
				Response.Write " (" & FormatNumber(GetSessao(global_cookie,"nrows_auts"&i),0) & " Reg.)"
			end if
		else
			Response.Write "<span class='span_imagem div_imagem_right icon_16 mozilla_blu' id='aba" & i & "_carregando'></span>"
			Response.Write " " & Servidores.ServList.Item(i-1).Nome
			Response.Write "<span id='aba" & i & "_total_resultados'></span>"
		end if

		Response.Write "			</a>"
		Response.Write "		</td>"
		
		if (i mod 5 = 0) AND (i <> iNumAba) then
			Response.Write "</tr>" 
			Response.Write "<tr>" 
		end if
	Next
	
	Response.Write "		<td class='td_right_bottom' style='background-color: #e9e9e9' colspan='" & (5 - (iNumAba mod 5)) & "'>"
	Response.Write "			&nbsp;"
	Response.Write "		</td>"
	Response.Write "	</tr>"
	Response.Write "	</table>"	
	Response.Write "  </td>"
	Response.Write "</tr>"	
	Response.Write "</table>"
	
	bCriouAbas = true
%>	
<script type="text/javascript"> 
	parent.hiddenFrame.iSrvCorrente_Aut = <%=iSrvCorrente%>;
</script>
<%				
end if

%>
<script type="text/javascript">
	parent.hiddenFrame.iNumAbas = <%=iNumAba%>;
</script>
<%
'******************************************************************************************

sPagCorrenteAbas = Request.QueryString("pagina")
if sPagCorrenteAbas = "" then	
	For iPag = 1 to Servidores.ServList.Count
	%>
		<script type="text/javascript">
			parent.hiddenFrame.arPagCorrenteAut[<%=iPag%>] = 1;
			parent.hiddenFrame.arIndCorrenteAut[<%=iPag%>] = 1;
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

'###################  DEBUG  ########################
'Response.Write "<br />Pagina antes: "&sPagCorrenteAbas
'Response.Write "<br />Indice antes: "&sIndCorrenteAbas
'Response.Write "<br />"
'####################################################

'Replica as ações para cada servidor
for iSrv = 1 to iNumAba

	'Atualiza o valor de iIndexSrv quando pesquisar por todos os servidores
	if iNumAba <> 1 then
		iIndexSrv = iSrv
	else
		if config_multi_servbib = 0 then
			iIndexSrv = 1
		else
			iIndexSrv = iSrvCorrente
		end if
	end if
	
	'Define a visibilidade dos Div's
	if iSrv = cInt(iSrvCorrente) or iNumAba = 1 then
		sDisplay = "block"
	else
		sDisplay = "none"
	end if

	sXMLFichas = ""
	
	' Marca se a pesquisa retornou algum resultado
	bResultOk = True
	bExibeFicha = True
	
	'Variáveis utilizadas na chamada da rotina "Paginacao"	
	pagina = arPaginas(iIndexSrv)
	indice = arIndices(iIndexSrv)
	
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	
	if len(trim(Request.QueryString("cod"))) > 0 AND Request.QueryString("veio_de") = "linkAutInfo" then
		aut_pag = Request.QueryString("cod")
	elseif Request.Form("submeteu") = "aut" AND Request.QueryString("veio_de") <> "paginacao" then
		bExibeFicha = False
		modo_busca = "rapida"
		%> 
			<script type="text/javascript">
				addBuscaEvent(SubmeteBuscaFrame, 4, <%=iIndexSrv%>);
			</script>
		<%
	else
		aut_pag = right(Request.QueryString("aut_pag"),len(Request.QueryString("aut_pag"))-1)
		aut_pag = left(aut_pag,len(aut_pag)-1)
		
		'###################  DEBUG  ########################
		'Response.Write "(ASP) vet_pag: "&aut_pag
		'####################################################
				
		arPagSrv = Split(aut_pag, "|")	
		aut_pag = arPagSrv(iIndexSrv)
		
		'###################  DEBUG  ########################
		'Response.Write "<br /> (ASP) arPagSrv("&iSrv&"): "&arPagSrv(iSrv)
		'####################################################
				
		if (aut_pag = "0") OR (aut_pag = "") then
			sMsgResult = "<div class='centro'><span style='color: red'>"&getTermo(global_idioma, 1341, "Nenhum registro encontrado.", 0)&"</div>"
			bResultOk = False
		end if
	end if
	
	'******************************************************************************************
	if sXMLFichas = "" then
	
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
	
		Set ROService = ROServer.CreateService("Web_Consulta")
		sXMLFichas = ROService.MontaFichasAutoridade(aut_pag,global_idioma, objParamDestaca)
		Set ROService = nothing
	end if
	'******************************************************************************************
	
	'Cria div de índice iSrv
	Response.Write "<div id='a_div_aba"&iIndexSrv&"' style='display:"&sDisplay&"; min-height: 288px;'>"
	
	if bExibeFicha then
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
	else
		Response.Write "<div id='a_div_aba"&iIndexSrv&"_resultado' style='min-height: 288px;'>"
		
		if (iNumAba > 1) then
			Response.Write "<br />"&getTermo(global_idioma,32,"Carregando",0)&"..."
		else
			Response.Write "<span class='span_imagem div_imagem_right icon_16 mozilla_blu '></span><br /><br /><br />"&getTermo(global_idioma,32,"Carregando",0)&"..."
		end if
		Response.Write "</div>"
	end if
	
	Response.Write "</div>"
Next
%>
</tr>
</table>
</td>
</tr>
</table>