<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top" style="display: block">
<script type='text/javascript' src="scripts/ajxSeries.js"></script>
<%

if Session("codigo_usuario") = "" then
	Response.Redirect("asp/logout.asp?logout=sim&modo_busca=" & GetModo_Busca)
end if

if Session("nome_usuario") <> "" then

	if config_multi_servbib = 1 then
		iIndexSrv = Session("Servidor_Logado")

		if iIndexSrv = "" then
			iIndexSrv = 1
		end if

		'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
		%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	end if

	Nome_Usuario = Session("nome_usuario")
	
	'Indica se a bibliografia a ser exibida é para ser do usuário logado OU de uma pesquisa previamente realizada
	if (Request.QueryString("curso") <> "0") then
		curso = Request.QueryString("curso")
		serie = Request.QueryString("serie")
		disciplina = Request.QueryString("disciplina")
		codigo_usu = "0"
	else
		curso = "0"
		serie = "0"
		disciplina = "0"
		codigo_usu = Session("codigo_usuario")
	end if
	
	Set ROService = ROServer.CreateService("Web_Consulta")
	Set ObjBibCurso = ROServer.CreateComplexType("TBibCurso")
	Set ObjBibCurso = ROService.BibliografiaCurso(codigo_usu,curso,serie,disciplina,global_idioma)
	Resultado = ObjBibCurso.Resultados
	iCurso = ObjBibCurso.iCurso
	iSerie = ObjBibCurso.iSerie
	iDisciplina = ObjBibCurso.iDisciplina
	sXMLPaginas = ObjBibCurso.Paginacao
	
	'Monta a informação acadêmica do usuário logado OU a escolhida pelo usuário
	sCurso = ROService.MontaInformacaoAcademica(iCurso, iSerie, iDisciplina)
	
	Set ObjBibCurso = nothing
	Set ROService = nothing
	
	'INÍCIO da listagem de cursos, séries e disciplinas
	if (((global_versao = vSOPHIA) and (global_academico = 1)) or ((global_numero_serie = 4184) or (global_numero_serie = 4090))) then %>
		
		<form name="frm_bib_curso" method="post" action="bib_curso.asp" onsubmit="ValidaBuscaAcademica();">
			
			<script type="text/javascript">
			    function atualiza_combo(serie_default) {	
			        atualizaSerie(<%=iIndexSrv%>,serie_default,document.frm_bib_curso);
			    }
				
			    function ValidaBuscaAcademica(){
			        //Validação do curso
			        if (document.frm_bib_curso.curso != null) {
			            if (document.frm_bib_curso.curso.selectedIndex < 1){
			                alert('<%=getTermo(global_idioma, 33, "É necessário selecionar o curso.", 0)%>');
			                return false;
			            }
			            else{
			                atualizaDadosAcademicosBibCurso();
			                LinkBibCurso(parent.hiddenFrame.modo_busca, 'visualizar');	
			            }
			        }
			    }
			</script>

			<br />
			<div id="academico" class="esquerda" style="margin-left: 10px;">
				<%
					Response.Write getTermo(global_idioma, 5914, "Bibliografia do curso", 0) & "&nbsp;"
					
					'INÍCIO da listagem do curso
					Response.Write "<select id='curso' name='curso' class='styled_combo' style='width: 300px;' onChange='atualiza_combo(-1);'>"
					Response.Write "<option value=-1>---------------------------</option>"

					'Carrega combo com cursos
					Set ROService = ROServer.CreateService("Web_Consulta")	
					sXML = ROService.DadosTabela(roTAB_CURSOS,"BIB_CURSO",0,repositorio_institucional)
					Set ROService = nothing
					TrataErros(1)

					'Indica se o usuário tem um curso pré-definido
					bTemCurso = false
					
					if (left(sXML,5) = "<?xml") then
						'Processa o XML
						Set xmlDoc = CreateObject("Microsoft.xmldom")
						xmlDoc.async = False
						xmlDoc.loadxml sXML
						Set xmlRoot = xmlDoc.documentElement

						'Verifica se a tabela foi encontrada e se possui registros
						if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
							For Each xmlTabela In xmlRoot.childNodes				
								if xmlTabela.nodeName  = "REGISTRO" then							
									sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
									sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
										
									if cLng(iCurso) = cLng(sCod) then
										selected = "selected"
										bTemCurso = true
									else
										selected = ""
									end if
									Response.Write "<option value="&sCod&" title='"&sDesc&"' "&selected&">"&sDesc&"</option>"										
								end if
							Next
						end if
					end if					
					Response.Write "</select>"
					'FIM da listagem do curso
					
					'Seleciona somente as séries do curso
					if (bTemCurso) then
						curso_atual = iCurso
					else
						curso_atual = 0
					end if
					
					'INÍCIO da listagem de série
					if ((global_numero_serie <> 4184) and (global_numero_serie <> 4090)) then
						Response.Write "&nbsp;&nbsp;"
						Response.Write "<select id='serie' name='serie' class='styled_combo' style='width: 100px;'>"
						Response.Write "<option value=-1>" & getTermo(global_idioma, 3827, "Todas", 0) & "</option>"
						Set ROService = ROServer.CreateService("Web_Consulta")	
						sXML = ROService.DadosSerieCurso(curso_atual)
						Set ROService = nothing
						TrataErros(1)

						if (left(sXML,5) = "<?xml") then
							'Processa o XML
							Set xmlDoc = CreateObject("Microsoft.xmldom")
							xmlDoc.async = False
							xmlDoc.loadxml sXML
							Set xmlRoot = xmlDoc.documentElement

							'Verifica se a tabela foi encontrada e se possui registros
							if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
								For Each xmlTabela In xmlRoot.childNodes				
									if xmlTabela.nodeName  = "REGISTRO" then							
										sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
										sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
											
										if cLng(iSerie) = cLng(sCod) then
											selected = "selected"
										else
											selected = ""
										end if
										Response.Write "<option value="&sCod&" title='"&sDesc&"' "&selected&">"&sDesc&"</option>"										
									end if
								Next
							end if
						end if					
						Response.Write "</select>"
					end if
					'FIM da listagem de série
					
					'INÍCIO da listagem de disciplina
					Response.Write "&nbsp;&nbsp;"
					Response.Write "<select id='disciplina' name='disciplina' class='styled_combo' style='width: 150px;'>"
					Response.Write "<option value=-1>" & getTermo(global_idioma, 3827, "Todas", 0) & "</option>"
					'Carrega combo com disciplinas
					Set ROService = ROServer.CreateService("Web_Consulta")	
					sXML = ROService.DadosTabela(roTAB_DISCIPLINAS,"",0,repositorio_institucional)
					Response.write sXML
					Set ROService = nothing
					TrataErros(1)

					if (left(sXML,5) = "<?xml") then
						'Processa o XML
						Set xmlDoc = CreateObject("Microsoft.xmldom")
						xmlDoc.async = False
						xmlDoc.loadxml sXML
						Set xmlRoot = xmlDoc.documentElement

						'Verifica se a tabela foi encontrada e se possui registros
						if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
							For Each xmlTabela In xmlRoot.childNodes				
								if xmlTabela.nodeName  = "REGISTRO" then							
									sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
									sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
									
									if cLng(iDisciplina) = cLng(sCod) then
										selected = "selected"
									else
										selected = ""
									end if
									Response.Write "<option value="&sCod&" title='"&sDesc&"' "&selected&">"&sDesc&"</option>"										
								end if
							Next
						end if
					end if					

					Response.Write "</select>"
					'FIM da listagem de disciplina
					
					Response.Write "&nbsp;&nbsp;"
				%>
				<input class="button_busca" type="button" title="<%=getTermo(global_idioma, 12, "Visualizar", 0)%>" value="<%=getTermo(global_idioma, 12, "Visualizar", 0)%>" onclick="ValidaBuscaAcademica();">
			</div>
		</form>
	
<%	
		'Caso o usuário clique no detalhe de um material, sem alterar as informações acadêmicas
		if (not Request.QueryString("curso") <> "0") then
			Response.Write "<script type='text/javascript'>if ((document.frm_bib_curso.curso != null) && (document.frm_bib_curso.curso.selectedIndex > 0)) ValidaBuscaAcademica();</script>"
		end if
	end if
	'FIM da listagem de cursos, séries e disciplinas
	
	if (len(Request.QueryString("vetor_pag")) > 0) then
		vet_pag = right(Request.QueryString("vetor_pag"),len(Request.QueryString("vetor_pag"))-1)
		vet_pag = left(vet_pag,len(vet_pag)-1)
	else
		vet_pag = ""
	end if	
	
	'###################  DEBUG  ########################
	'Response.Write "(ASP) vet_pag: "&vet_pag
	'####################################################
	
	arPagSrv = Split(vet_pag, "|")	
	vet_pag = arPagSrv(iIndexSrv)

	'###################  DEBUG  ########################
	'Response.Write "<br /> (ASP) arPagSrv("&iIndexSrv&"): "&arPagSrv(iIndexSrv)
	'####################################################
			
	'Utilizado para montar a paginação
	'******************************************************************************************
	%><!-- #include file="monta_vetores.asp" --><%
	'******************************************************************************************
	pagina = 0
	indice = 0
	if (vet_pag = "0") OR (vet_pag = "") OR (vet_pag = "undefined") then
		if (Request.QueryString("veio_de") = "detalhe") then
			vet_pag = GetSessao(global_cookie,"bibcurso")
			pagina = GetSessao(global_cookie,"bibcurso_pag")
			indice = GetSessao(global_cookie,"bibcurso_indice")
		else
			vet_pag = vetor_pag(1)
		end if
	end if
			
	'Destaque dos termos pesquisados
	Set objParamDestaca = ROServer.CreateComplexType("TParamBuscaHighlight")	
	
	redim array_campos (5)
	redim array_palavras (5)
	redim array_frase_exata (5)

	'Todos os campos
	array_campos(0) = 0
	array_palavras(0) = ""
	array_frase_exata(0) = ""
	
	'Autor
	array_campos(1) = 0
	array_palavras(1) = ""
	array_frase_exata(1) = ""
	
	'Assunto
	array_campos(2) = 0
	array_palavras(2) = ""
	array_frase_exata(2) = ""
	
	'Ementa
	array_campos(3) = 0
	array_palavras(3) = ""
	array_frase_exata(3) = ""
	
	'Campo vazio
	array_campos(4) = 0
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

	objParamDestaca.asPalavras     = aiPalavras
	objParamDestaca.asFraseExata   = aiFrase
	objParamDestaca.aiCampos       = aiCampos

	'Orgão de origem
	objParamDestaca.iOrgOrigem     =  -1

	'Norma
	objParamDestaca.iNorma     =  -1

	'Número
	objParamDestaca.sNumero        = ""
	
	'Fim do destaque dos termos pesquisados
	
	Set ROService = ROServer.CreateService("Web_Consulta")
    Set RetornoMontaFichas = ROService.MontaFichas(vet_pag, "", false, global_idioma, objParamDestaca)
	sXMLFichas = RetornoMontaFichas.sMsg
    Set RetornoMontaFichas = nothing
	Set ROService = nothing

	if (trim(sXMLFichas) <> "") then
		'Variáveis utilizadas na chamada da rotina "Paginacao"
		sPagCorrenteAbas = Request.QueryString("pagina")
		if sPagCorrenteAbas = "" then	
			%>
				<script type='text/javascript'>
				    parent.hiddenFrame.arPagCorrenteBibCurso[1] = 1;
				    parent.hiddenFrame.arIndCorrenteBibCurso[1] = 1;
				</script>
			<%
			sPagCorrenteAbas = "|1" + sPagCorrenteAbas
		end if

		sIndCorrenteAbas = Request.QueryString("indice")
		if sIndCorrenteAbas = "" then
			sIndCorrenteAbas = sPagCorrenteAbas
		end if

		arPaginas = Split(sPagCorrenteAbas, "|")
		arIndices = Split(sIndCorrenteAbas, "|")
		
		'Essas variáveis podem ter sido alteradas antes, se a página 
		'que requisitou a bib_curso foi a de um detalhe de material
		if (pagina <= 0) then
			pagina = arPaginas(iIndexSrv)
		end if
		
		if (indice <= 0) then
			indice = arIndices(iIndexSrv)
		end if
		sOrigem = "BibCurso"
		'Fim das Variáveis utilizadas na chamada da rotina "Paginacao"
	
		SalvaSessao "bibcurso",vet_pag,global_cookie
		SalvaSessao "bibcurso_pag",pagina,global_cookie
		SalvaSessao "bibcurso_indice",indice,global_cookie
			
		Response.Write "<br/>"
        Response.Write "<table class='tab_paginacao max_width' style='border-spacing: 1px; padding: 0'><tr>"
		Response.Write "<td class='centro' style='height: 20px'>"
		Response.Write "<b>"&getTermo(global_idioma, 1347, "Bibliografia do curso de", 0)&"&nbsp;" & sCurso & "</b>"
		Response.Write "</td></tr></table>"
		
		'Cria menu de ações (Selecionar, Marcar Todos, etc)
		Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
		Response.Write "	<tr style='height: 22px'>"
		Response.Write "		<td class='esquerda'>&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1351, "Clique aqui para selecionar todos os títulos desta página.", 0)&"' href=""javascript:fichas_marcar_todos("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-check'></span>"&getTermo(global_idioma, 1348, "Selecionar todos", 0)&"</a>&nbsp;&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1352, "Clique aqui para desmarcar todos os títulos desta página.", 0)&"' href=""javascript:fichas_desmarcar_todos("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-delete'></span> "&getTermo(global_idioma, 1349, "Desmarcar selecionados", 0)&"</a>&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1353, "Clique aqui para enviar os títulos selecionados para Minha seleção.", 0)&"' href=""javascript:enviar_minha_selecao_bibcurso("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-selecao'></span> "&getTermo(global_idioma, 1350, "Enviar para minha seleção", 0)&"&nbsp;&nbsp;"	
		Response.Write "            <a class='link_serv_custom' title='"&getTermo(global_idioma, 8318, "Clique aqui para salvar os títulos selecionados nos favoritos.", 0)&"' href=""javascript:salvar_favoritos("&iIndexSrv&",'resultado',true);"">"
		Response.Write "            <span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-star'></span>"&getTermo(global_idioma, 8317, "Salvar favoritos", 0)&"</a>"
		Response.Write "		</td>"	
		Response.Write "	</tr>"
		Response.Write "	</table>"
		
		'Cria tabela de paginação
		Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
		Response.Write "	<tr style='height: 26px'>"
		Response.Write "		<td class='esquerda'>&nbsp;&nbsp;"

		Paginacao GetSessao(global_cookie,"nrows_bib_curso"&iIndexSrv),global_num_linhas,global_max_links,pagina,indice,modo_busca,"bib_curso",GetSessao(global_cookie,"nrows_bib_curso_real"&iIndexSrv)

		Response.Write "		</td>"
		Response.Write "	</tr>"
		Response.Write "	</table>"
		'Fim de cria tabela de paginação
		
		%><!-- #include file='grid_ficha.asp' --><%	
		
		'Cria tabela de paginação
		Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
		Response.Write "	<tr style='height: 26px'>"
		Response.Write "		<td class='esquerda'>&nbsp;&nbsp;"

		Paginacao GetSessao(global_cookie,"nrows_bib_curso"&iIndexSrv),global_num_linhas,global_max_links,pagina,indice,modo_busca,"bib_curso",GetSessao(global_cookie,"nrows_bib_curso_real"&iIndexSrv)

		Response.Write "		</td>"
		Response.Write "	</tr>"
		Response.Write "	</table>"
		'Fim de cria tabela de paginação
		
		'Cria menu de ações (Selecionar, Marcar Todos, etc)
		Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
		Response.Write "	<tr style='height: 22px'>"
		Response.Write "		<td class='esquerda'>&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1351, "Clique aqui para selecionar todos os títulos desta página.", 0)&"' href=""javascript:fichas_marcar_todos("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-check'></span> "&getTermo(global_idioma, 1348, "Selecionar todos", 0)&"</a>&nbsp;&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1352, "Clique aqui para desmarcar todos os títulos desta página.", 0)&"' href=""javascript:fichas_desmarcar_todos("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-delete'></span> "&getTermo(global_idioma, 1349, "Desmarcar selecionados", 0)&"</a>&nbsp;&nbsp;"
		Response.Write "			<a class='link_serv_custom' title='"&getTermo(global_idioma, 1353, "Clique aqui para enviar os títulos selecionados para Minha seleção.", 0)&"' href=""javascript:enviar_minha_selecao_bibcurso("&iIndexSrv&");"">"
		Response.Write "			<span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-selecao'></span> "&getTermo(global_idioma, 1350, "Enviar para minha seleção", 0)&"&nbsp;&nbsp;"	
		Response.Write "            <a class='link_serv_custom' title='"&getTermo(global_idioma, 8318, "Clique aqui para salvar os títulos selecionados nos favoritos.", 0)&"' href=""javascript:salvar_favoritos("&iIndexSrv&",'resultado',true);"">"
		Response.Write "            <span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-star'></span>"&getTermo(global_idioma, 8317, "Salvar favoritos", 0)&"</a>"
		Response.Write "		</td>"	
		Response.Write "	</tr>"
		Response.Write "	</table>"
	else 
		Response.Write "<br /><br />"
		Response.Write getTermoHtml(global_idioma, 1346, "Não há títulos definidos para o curso.", 0)
	end if
	
else%>
	erro
<%end if%>
</td>
</tr>
</table>