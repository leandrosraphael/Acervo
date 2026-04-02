<%
	'--------------------------------------------------------------------------------
	' 				VERIFICA CURTIR DO FACEBOOK HABILITADO
	'--------------------------------------------------------------------------------

	if global_facebook_curtir = 1 then
		codigoIdioma = "pt_BR"
		if global_idioma = 1 then
			codigoIdioma = "es_ES"
		elseif global_idioma = 2 then
			codigoIdioma = "en_US"
		elseif global_idioma = 3 then
			codigoIdioma = "ca_ES"
		end if

		' Utilização do Curtir do Facebook via API	
		if global_facebook_iframe <> 1 then	
%>
			<div id="fb-root"></div>
			<script>
			    window.fbAsyncInit = function () {
			        FB.init({
			            appId: '<%=global_facebook_appid %>', // App ID
			            status: true, // check login status
			            xfbml: true  // parse XFBML
			        });

			        FB.Event.subscribe('xfbml.render', function (response) {
			            $(".p-facebook").css("visibility", "visible");
			            $(".p-facebook").css("overflow", "visible");
			            setTimeout(function () {
			                $(".p-facebook iframe").css("width", "100px");
			            }, 100);
			        });
			    };

			    // Load the SDK Asynchronously
			    (function (d) {
			        var js, id = 'facebook-jssdk', ref = d.getElementsByTagName('script')[0];
			        if (d.getElementById(id)) { return; }
			        js = d.createElement('script'); js.id = id; js.async = true;
			        js.src = "//connect.facebook.net/<%=codigoIdioma %>/all.js";
			        ref.parentNode.insertBefore(js, ref);
			    } (document));
		</script>
	<%	end if
	end if  %>

<script type="text/javascript">
function envia_combo(modo,cod_obra,qtde,pag,pos,servidor) {
	<% if Request.QueryString("veio_de") = "link_detalhe" then %>
		parent.mainFrame.location = 'index.asp?modo_busca='+modo+'&content=detalhe&codigo_obra='+cod_obra+'&tipo_obra=0&combo=1&qtde='+qtde+'&pagina='+pag+'&veio_de=link_detalhe&posicao_vetor='+pos+'&ano='+parent.mainFrame.document.frm_detalhe.ano.value+'&biblioteca='+parent.hiddenFrame.geral_bib+'&iBanner='+parent.hiddenFrame.iBanner + '&iIdioma='+parent.hiddenFrame.iIdioma+'&Servidor='+servidor;
	<% else %>
		parent.mainFrame.location = 'index.asp?modo_busca='+modo+'&content=detalhe&codigo_obra='+cod_obra+'&tipo_obra=0&combo=1&qtde='+qtde+'&pagina='+pag+'&veio_de=busca_principal&posicao_vetor='+pos+'&ano='+parent.mainFrame.document.frm_detalhe.ano.value+'&biblioteca='+parent.hiddenFrame.geral_bib+'&iBanner='+parent.hiddenFrame.iBanner + '&iIdioma='+parent.hiddenFrame.iIdioma+'&Servidor='+servidor;
	<% end if %>
}
</script>
<%
levantamento_bib = CLng(request.QueryString("levantamento"))
pagina 		  = Request.QueryString("pagina")
codigo 		  = Request.QueryString("codigo_obra") 
tipo_obra 	  = Request.QueryString("tipo_obra") 
tipo_artigo   = Request.QueryString("tipo")
filtro_artigo = Request.QueryString("filtra_artigo")
veio_de = Request.QueryString("veio_de")

bFiltroPorBiblioteca = (Trim(Request.QueryString("biblioteca") <> ""))
bTermoBibSingular = (InStr(Trim(Request.QueryString("biblioteca")), ",") = 0)

codigo_obra   = codigo

ScriptIntegra = ""

'Define o Servbib padrão da obra
iIndexSrv = IntQueryString("Servidor", 1)
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

'************************************************************
' FAZ A CHAMADA DA ROTINA QUE MONTA A FICHA SOPHIA
'************************************************************
On Error Resume Next
SET ROService = ROServer.CreateService("Web_Consulta")
if tipo_obra <> "" then
	'Destaque dos termos pesquisados
	redim array_campos (5)
	redim array_palavras (5)
	redim array_frase_exata (5)
	
	Set objParamDestaca = ROServer.CreateComplexType("TParamBuscaHighlight")
	
	modo_busca = GetModo_Busca
	
	if ((modo_busca = "rapida") or (modo_busca = "combinada")) then
	
		array_campos(0) = GetPosCampoPesquisa(Request.QueryString("campo1"))
		array_campos(1) = GetPosCampoPesquisa(Request.QueryString("campo2"))
		array_campos(2) = GetPosCampoPesquisa(Request.QueryString("campo3"))
		array_campos(3) = GetPosCampoPesquisa(Request.QueryString("campo4"))
		array_campos(4) = GetPosCampoPesquisa(Request.QueryString("campo5"))

		sValor = RemoveUnderline(Request.QueryString("valor1"))
		array_palavras(0) = SemAspas(sValor)
		array_frase_exata(0) = EntreAspas(sValor)
		
		sValor = RemoveUnderline(Request.QueryString("valor2"))
		array_palavras(1) = SemAspas(sValor)
		array_frase_exata(1) = EntreAspas(sValor)
		
		sValor = RemoveUnderline(Request.QueryString("valor3"))
		array_palavras(2) = SemAspas(sValor)
		array_frase_exata(2) = EntreAspas(sValor)
		
		sValor = RemoveUnderline(Request.QueryString("valor4"))
		array_palavras(3) = SemAspas(sValor)
		array_frase_exata(3) = EntreAspas(sValor)
		
		sValor = RemoveUnderline(Request.QueryString("valor5"))
		array_palavras(4) = SemAspas(sValor)
		array_frase_exata(4) = EntreAspas(sValor)
		
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

		if(Request.QueryString("meiofisico") <> "" and Request.QueryString("meiofisico") <> "0") then
			objParamDestaca.sMeioFisico  = Request.QueryString("meiofisico")
		end if

		if(Request.QueryString("niveis") <> "" and Request.QueryString("niveis") <> "0") then
			objParamDestaca.sNiveis  = Request.QueryString("niveis")
		end if

		if(Request.QueryString("formaregistro") <> "" and Request.QueryString("formaregistro") <> "0") then
			objParamDestaca.sFormaRegistro  = Request.QueryString("formaregistro")
		end if
	elseif (modo_busca = "legislacao") then
		'Campo + Valor para palavras + Valor para frase exata (sempre vazio)
	
		'Palavra-Chave
		array_campos(0) = GetPosCampoPesquisa(Request.QueryString("campo1"))
		sValor = RemoveUnderline(Request.QueryString("valor1"))
		array_palavras(0) = SemAspas(sValor)
		array_frase_exata(0) = EntreAspas(sValor)
			
		'Autor
		array_campos(1) = GetPosCampoPesquisa(Request.QueryString("campo2"))
		sValor = RemoveUnderline(Request.QueryString("valor2"))
		array_palavras(1) = SemAspas(sValor)
		array_frase_exata(1) = EntreAspas(sValor)

		'Assunto
		array_campos(2) = GetPosCampoPesquisa(Request.QueryString("campo4"))
		sValor = RemoveUnderline(Request.QueryString("valor4"))
		array_palavras(2) = SemAspas(sValor)
		array_frase_exata(2) = EntreAspas(sValor)
		
		'Ementa
		array_campos(3) = GetPosCampoPesquisa(Request.QueryString("campo5"))
		sValor = RemoveUnderline(Request.QueryString("valor5"))
		array_palavras(3) = SemAspas(sValor)
		array_frase_exata(3) = EntreAspas(sValor)

		'Texto integral
		array_campos(4) = GetPosCampoPesquisa(Request.QueryString("campo6"))
		sValor = RemoveUnderline(Request.QueryString("valor6"))
		array_palavras(4) = SemAspas(sValor)
		array_frase_exata(4) = EntreAspas(sValor)
			
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
		
		'Orgão de origem
		objParamDestaca.iOrgOrigem     = Request.QueryString("valor3")
		
		'Número
		objParamDestaca.sNumero        = RemoveUnderline(Request.QueryString("valor7"))
		
		'Norma
		objParamDestaca.iNorma         = Request.QueryString("valor8")
		
		'Lista Orgão de origem
		objParamDestaca.ListaOrgaoOrigem   =  Request.QueryString("orgao_origem")

		'Lista Normas
		objParamDestaca.ListaNormas     =  Request.QueryString("norma")

		'Autoria de projeto de lei
		objParamDestaca.sAutoriaProjetoLei = RemoveUnderline(Request.QueryString("valor9"))
		
		'Número de projeto de lei
		objParamDestaca.sNumeroProjetoLei = RemoveUnderline(Request.QueryString("valor10"))

        'Processo
		objParamDestaca.sProcesso = RemoveUnderline(Request.QueryString("valor15"))
	
	end if
	'Fim do destaque os termos pesquisados

	Set ROService = ROServer.CreateService("Web_Consulta")
	xml_ficha = ROService.MontaFichaSophiA(codigo, tipo_obra, global_idioma, objParamDestaca)

	if (trim(Session("codigo_usuario")) = "") then
		codigoUsuario = 0
	else
		codigoUsuario = ClNg(Session("codigo_usuario"))
	end if

    ROService.ContaAcessosTitulos codigo, codigoUsuario
	Tipo 	  = tipo_obra
	if filtro_artigo > 0 then
		if Tipo = 0 OR Tipo = 1 then
			tipo_artigo = 3
		else
			tipo_artigo = 0
		end if
	else
		tipo_artigo = 0
	end if
end if

TrataErros(1)

if (trim(Request.QueryString("ano")) = "") then
	cbAnoAtual = ""
else
	cbAnoAtual = trim(Request.QueryString("ano"))
end if

if len(trim(Session("codigo_usuario"))) = 0 then
	iUsuarioLogado = 0
else
	if config_multi_servbib = 1 then
		if CStr(Session("Servidor_Logado")) = CStr(iIndexSrv) then
			iUsuarioLogado = Session("codigo_usuario")
		else
			iUsuarioLogado = 0
		end if
	else
		iUsuarioLogado = Session("codigo_usuario")
	end if
end if

'************************************************************
' FAZ A CHAMADA DA ROTINA QUE MONTA OS EXEMPLARES
'************************************************************
ExibeEx = true

if tipo_artigo > 0 then
	'ANALITICA/ARTIGO
	if (CStr(Tipo) = "3") then
		ExibeEx = false
		xml_exemplar = ""
	else 
		xml_exemplar = ROService.MontaExemplar(codigo, tipo_artigo, filtro_artigo, 0, 0, cbAnoAtual, iUsuarioLogado, global_idioma, (global_IP_Local = 1))
	end if
else
	'ANALITICA/ARTIGO
	if (CStr(Tipo) = "3") then
		ExibeEx = false
		xml_exemplar = ""
	else 
		if CStr(Tipo) = "0" OR CStr(Tipo) = "1" then
			BIBLIOTECA = request.QueryString("biblioteca")
			PROJETO = request.QueryString("projeto")
		else
			BIBLIOTECA = ""
			PROJETO = 0
		end if
	
		'Quando o cliente utilizar a opção de múltiplos servidores, a biblioteca não deve ser passada como parâmetro
		if (config_multi_servbib = 1) then
			BIBLIOTECA = ""
		end if
		
		xml_exemplar = ROService.MontaExemplar(codigo, Tipo, 0, BIBLIOTECA, PROJETO, cbAnoAtual, iUsuarioLogado, global_idioma, (global_IP_Local = 1))
	end if
end if

TrataErros(1)

SET ROService = nothing

%>
<table class="max_width max_height">
<tr>
<td class="td_center_top td_padrao">

<%
'************************************************************
' MONTA NAVEGADORES
'************************************************************
Response.Write "<table class='removerBordas remover_bordas_padding tab_paginacao max_width'>"
Response.Write "<tr style='height: 26px'>"
Response.Write "<td style='width: 33%' class='esquerda'>"

sImgVoltar = "voltar"
if Request.QueryString("veio_de") = "circulacao" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkCirculacoes(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "reserva" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkReservas(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "bibcurso" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkBibCurso(parent.hiddenFrame.modo_busca,'detalhe');""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "sels" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkSelecao(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "leg" OR Request.QueryString("veio_de") = "analitica" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:history.go(-1);""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "favoritos" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkFavoritos(parent.hiddenFrame.modo_busca, "&Request.QueryString("listaSelecionada")&");""><span class='transparent-icon removerBordas span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
elseif Request.QueryString("veio_de") = "link_detalhe" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
elseif Request.QueryString("veio_de") = "ultimas_aquisicoes_home" then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:volta_resultado('ultimas_aquisicoes_home',0);""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
else
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:volta_resultado(parent.hiddenFrame.modo_busca,"&Request.QueryString("pagina")&");""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
end if
Response.Write "</td>"
Response.Write "<td class='td_center_middle' style='width: 34%;'>" <!--Barra de navegação - Página 1 de x-->
if Request.QueryString("veio_de") = "sels" then
	%><!-- #include file="navegador_sels.asp" --><%
elseif Request.QueryString("veio_de") = "reserva" OR Request.QueryString("veio_de") = "leg" OR Request.QueryString("veio_de") = "periodico" OR _
	   Request.QueryString("veio_de") = "link_detalhe" OR Request.QueryString("veio_de") = "analitica" then
	Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
else
	%><!-- #include file="navegador_detalhes.asp" --><%
end if
Response.Write "</td>"
Response.Write "<td class='direita' style='width: 33%'>" <!--Barra de navegação - "Nova pesquisa"-->
if (somente_detalhe = false) then
	Response.Write "<a class='link_serv_custom' title='"&getTermo(global_idioma, 1345, "Clique aqui para fazer uma nova consulta.", 0)&"' href=""javascript:novaPesquisa(parent.hiddenFrame.modo_busca);"" onClick=Resetar();>"
	Response.Write "<span class='transparent-icon span_imagem icon_16 icon-small-newsearch'></span>"
	Response.Write "&nbsp;"&getTermo(global_idioma, 1344, "Nova pesquisa", 0)&"</a>"
end if
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
Response.Write "</tr></table>"

'************************************************************
' INICIO DO DETALHE
'************************************************************
if cStr(Tipo) = "1" then
	desc_tipo_det = getTermoHtml(global_idioma, 1544, "Detalhes da obra", 0)
	desc_tipo_ficha = getTermoHtml(global_idioma, 1545, "Ficha resumida da obra", 0)
elseif cStr(Tipo) = "0" then
	desc_tipo_det = getTermo(global_idioma, 1546, "Detalhes do periódico", 0)
	desc_tipo_ficha = getTermo(global_idioma, 1547, "Ficha resumida do periódico", 0)
elseif cStr(Tipo) = "3" then
	desc_tipo_det = getTermoHtml(global_idioma, 1548, "Detalhes da analítica", 0)
	desc_tipo_ficha = getTermoHtml(global_idioma, 1549, "Ficha resumida da analítica", 0)
elseif cStr(Tipo) = "2" then
	if (global_numero_serie = 3179) then
		desc_tipo_det = "Detalhes do ato normativo"
		desc_tipo_ficha = "Ficha resumida do ato normativo"
	else
		desc_tipo_det = getTermo(global_idioma, 1550, "Detalhes da legislação", 0)
		desc_tipo_ficha = getTermo(global_idioma, 1551, "Ficha resumida da legislação", 0)
	end if
elseif cStr(Tipo) = "20" then
	desc_tipo_det = getTermo(global_idioma, 1552, "Detalhes da fonte avulsa de periódico", 0)
	desc_tipo_ficha = getTermo(global_idioma, 1553, "Ficha resumida da fonte avulsa de periódico", 0)
elseif cStr(Tipo) = "21" then
	desc_tipo_det = getTermo(global_idioma, 1554, "Detalhes da fonte avulsa de obra", 0)
	desc_tipo_ficha = getTermo(global_idioma, 1555, "Ficha resumida da fonte avulsa de obra", 0)
end if

'BND
'if (global_numero_serie = 5516) then
	desc_tipo_ficha = desc_tipo_det
'end if

if Request.QueryString("veio_de") = "bibcurso" then
	'Monta a informação acadêmica do usuário logado OU a escolhida pelo usuário
	curso = Request.QueryString("curso")
	serie = Request.QueryString("serie")
	disciplina = Request.QueryString("disciplina")

	Set ROService = ROServer.CreateService("Web_Consulta")
	sDadosAcademicos = ROService.MontaInformacaoAcademica(curso, serie, disciplina)
	Set ROService = nothing
	
	if (sDadosAcademicos <> "") then
		desc_tipo_ficha = desc_tipo_ficha & " - " & getTermo(global_idioma, 5914, "Bibliografia do curso", 0) & " " & sDadosAcademicos
	end if
end if

'************************************************************
' MONTA AS ABAS (FICHA RESUMIDA / FICHA COMPLETA / MARC)
'************************************************************
if cInt(global_marc) = 1 AND cInt(Tipo) <> 2 AND cInt(Tipo) >= 0 AND cInt(Tipo) <= 3 then%>
	<table class="remover_bordas_padding max_width" style="table-layout: fixed; border-color: #CCCCCC">
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
			else
				FichaResumida = getTermo(global_idioma, 1032, "Detalhes", 0)
				ocultaItem = ""
			end if
		%>

		<td class="td_center_middle td_abas_detalhe_aacr2" id="tdx_a1"><a class="link_abas" href="#" onclick="habilitaMarc(0,parent.hiddenFrame.layerX,'detalhes',<%=codigo_obra%>,<%=Tipo%>);" id="lk_a1"><%=FichaResumida%></a></td>
		<td class="td_center_middle td_abas_detalhe_marc" id="tdx_b1"><a class="link_abas" href="#" onclick="habilitaMarc(1,parent.hiddenFrame.layerX,'marcTags',<%=codigo_obra%>,<%=Tipo%>);" id="lk_b1"><%=getTermo(global_idioma, 997, "MARC tags", 0)%></a></td>
		<td class="td_center_middle td_abas_detalhe_marc" id="tdx_dc1"><a class="link_abas" href="#" onclick="habilitaMarc(1,parent.hiddenFrame.layerX,'dublinCore',<%=codigo_obra%>,<%=Tipo%>);" id="lk_dc1">Dublin Core</a></td>


		<% if (possui_estrutura_hierarquica) then %>
			<td class="td_center_middle td_abas_detalhe_marc" id="tdx_c1"><a class="link_abas" href="#" onclick="habilitaMarc(1,parent.hiddenFrame.layerX,'arquivo',<%=codigo_obra%>,<%=Tipo%>);" id="lk_c1"><%=getTermo(global_idioma, 7482, "Registros relacionados", 0)%></a></td>
		<% end if %>

		<td class="td_right_bottom">&nbsp;</td>
	</tr>
	<tr>
		<td class="td_right_bottom td_abas_detalhe_aacr2_esquerda" id="tdx_dt1">&nbsp;</td>
		<td class="td_right_bottom td_abas_detalhe_aacr2_centro" id="tdx_dt2">&nbsp;</td>

		<% if (possui_estrutura_hierarquica) then %>
			<td class="td_right_bottom td_abas_detalhe_aacr2_centro" id="tdx_dt3">&nbsp;</td>
			<td class="td_right_bottom td_abas_detalhe_aacr2_centro" id="tdx_dt5">&nbsp;</td>
			<td class="td_right_bottom td_abas_detalhe_aacr2_direita" id="tdx_dt4">&nbsp;</td>
		<% else %>
			<td class="td_right_bottom td_abas_detalhe_aacr2_centro" id="tdx_dt5">&nbsp;</td>
			<td colspan="2" class="td_right_bottom td_abas_detalhe_aacr2_direita" id="tdx_dt4">&nbsp;</td>
		<% end if %>
	</tr>
	<tr>
		<td colspan="5" class='td_center_middle ficha_detalhes'>&nbsp;<b><%=desc_tipo_ficha%></b>&nbsp;</td>
	</tr>
	<tr><td colspan="5" class="td_center_top td_marc"><br/>
<% else %>
	<table class="remover_bordas_padding max_width" style="border-color: #CCCCCC"><tr>
	<td class="td_center_top">&nbsp;</td></tr><tr>
	<% if cStr(Tipo) = "2" then %>
		<td class='td_center_middle ficha_legislacao'>&nbsp;<b><%=desc_tipo_det%></b>&nbsp;</td></tr>
	<% else %>	
		<td class='td_center_middle ficha_detalhes'>&nbsp;<b><%=desc_tipo_det%></b>&nbsp;</td></tr>
	<%end if%>	
	<tr><td class="td_center_top td_marc"><br/> <!--Detalhes da legislação--> 
<% end if 
Response.write "<table class='remover_bordas_padding max_height' style='width: 100%; display: table; padding: 1%;'>"
Response.Write "<tr><td>"

'*************************************
'Serviços
'*************************************
serv_selecao   = false
serv_reserva   = false
serv_aquisicao = false
serv_midias    = false
serv_refbib    = false
serv_bibcomp   = false
serv_analitica = false

'**************************************************************************
' MONTA O DETALHE A PARTIR DO XML
'**************************************************************************
if left(xml_ficha,5) = "<?xml" then
	
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml xml_ficha
	Set xmlRoot = xmlDoc.documentElement
	
	'*************************************
	'EXTRAÇÃO DE TÍTULO E AUTOR, QUANDO NÃO EXISTIR CAPA PARA O ELEMENTO PESQUISADO
	'*************************************
	titulo_autor_sem_capa = "<div class='div_capa'>"
	titulo_sem_capa = ""
	desc_autor_sem_capa = ""
	complemento_sem_capa = ""
	url_imagem_referencia = ""
	codigo_material = 0

	For Each xmlPNome In xmlRoot.childNodes 
        
        if xmlPNome.nodeName  = "IDENTIFICADORES" then             
            For Each xmlIdentificadores in xmlPNome.childNodes
                tipoIdentificador = xmlIdentificadores.attributes.getNamedItem("Tipo").value
                imagemIdentificador = "data:image;base64," & xmlIdentificadores.attributes.getNamedItem("Base64").value
                %><script>sessionStorage.setItem("<%=tipoIdentificador%>" ,"<%=imagemIdentificador%>");</script><%                    
            next
        end if

		if xmlPNome.nodeName = "FICHA" then
     
			tem_capa = xmlPNome.attributes.getNamedItem("CAPA").value

			For Each xmlCampos In xmlPNome.childNodes

				if xmlCampos.nodeName = "TITULO" OR xmlCampos.nodeName = "NORMA" then

					if(xmlCampos.nodeName = "TITULO") then
						titulo_sem_capa = xmlCampos.attributes.getNamedItem("Valor").value
						titulo_f_sem_capa = Replace(Replace(titulo_sem_capa,"#D#",""),"#/D#","")
						complemento_sem_capa = " " & Replace(Replace(xmlCampos.attributes.getNamedItem("Complemento").value,"#D#",""),"#/D#","")
						titulo_f_sem_capa = titulo_f_sem_capa & complemento_sem_capa
						if( Len(titulo_f_sem_capa) >= 45) then
							titulo_f_sem_capa = left(titulo_f_sem_capa,45) & "..."
						end if
					else
						titulo_sem_capa = xmlCampos.attributes.getNamedItem("Valor").value
						titulo_f_sem_capa = Replace(Replace(titulo_sem_capa,"#D#",""),"#/D#","")
						titulo_f_sem_capa = titulo_f_sem_capa & complemento_sem_capa
						if( Len(titulo_f_sem_capa) >= 45) then
							titulo_f_sem_capa = left(titulo_f_sem_capa,45) & "..."
						end if
					end if

				elseif xmlCampos.nodeName = "ENT_PRINC" OR xmlCampos.nodeName = "ORGAO_ORIGEM" then

					desc_autor_sem_capa = xmlCampos.attributes.getNamedItem("Valor").value 
					desc_autor_sem_capa = Replace(Replace(desc_autor_sem_capa,"#D#",""),"#/D#","")
					if( Len(desc_autor_sem_capa) >= 25 ) then
						desc_autor_sem_capa = left(desc_autor_sem_capa,25) & "..."
					end if
					desc_autor_sem_capa = "<br/><br/>" & "<span>" & desc_autor_sem_capa & "</span>"

				' Adequação BND - A imagem de referência irá remeter para o site cadastrado
				elseif xmlCampos.nodeName = "SITES" AND global_numero_serie = 5516 then

					For Each xmlSubCampos In xmlCampos.childNodes
                        For Each xmlItem In xmlSubCampos.childNodes
                            codMidia = xmlItem.attributes.getNamedItem("Codigo").value
                            dspace = xmlItem.attributes.getNamedItem("DSpace").value
						    url_site = TrocaTagMarcador(xmlItem.attributes.getNamedItem("Valor").value)
						    if (UCase(Right(url_site, 3)) = "JPG") AND url_imagem_referencia = "" then
							    url_imagem_referencia = url_site
						    end if
                        Next
					Next
				elseif xmlCampos.nodeName = "INF_PUBLICACAO" AND global_numero_serie = 5516 then
					codigo_material = xmlCampos.attributes.getNamedItem("Codigo").value
				end if

			next

		end if        
	next

	' Adequação BND - A imagem de referência irá remeter para o site cadastrado
	if global_numero_serie = 5516 AND url_imagem_referencia = "" then
		For Each xmlPNome In xmlRoot.childNodes
			if xmlPNome.nodeName = "FICHA" then
				For Each xmlCampos In xmlPNome.childNodes
					if xmlCampos.nodeName = "SITES" AND url_imagem_referencia = "" then
						For Each xmlSubCampos In xmlCampos.childNodes
                            For Each xmlItem In xmlSubCampos.childNodes
                                codMidia = xmlItem.attributes.getNamedItem("Codigo").value
                                dspace = xmlItem.attributes.getNamedItem("DSpace").value
							    url_site = TrocaTagMarcador(xmlItem.attributes.getNamedItem("Valor").value)
							    if (UCase(Right(url_site, 3)) = "PDF") AND url_imagem_referencia = "" then
								    url_imagem_referencia = url_site
    							end if
                            Next
						Next
					end if
				Next
				if url_imagem_referencia = "" then
					For Each xmlCampos In xmlFicha.childNodes
						if xmlCampos.nodeName = "SITES" AND url_imagem_referencia = "" then
                            For Each xmlSubCampos In xmlCampos.childNodes
                                For Each xmlItem In xmlSubCampos.childNodes
                                    codMidia = xmlItem.attributes.getNamedItem("Codigo").value
                                    dspace = xmlItem.attributes.getNamedItem("DSpace").value
							        url_site = RemoveTagMarcador(xmlItem.attributes.getNamedItem("Valor").value)
							        if (UCase(Right(url_site, 3)) = "MP3") AND url_imagem_referencia = "" then
								        url_imagem_referencia = url_site
								        capa_audio = true
    							    end if
                                Next
                            Next
						end if
					Next
				end if
				if url_imagem_referencia = "" then
					For Each xmlCampos In xmlFicha.childNodes
						if xmlCampos.nodeName = "SITES" AND url_imagem_referencia = "" then
                            For Each xmlSubCampos In xmlCampos.childNodes
                                For Each xmlItem In xmlSubCampos.childNodes
                                    codMidia = xmlItem.attributes.getNamedItem("Codigo").value
                                    dspace = xmlItem.attributes.getNamedItem("DSpace").value
							        url_site = RemoveTagMarcador(xmlItem.attributes.getNamedItem("Valor").value)
							        if (UCase(Right(url_site, 3)) = "MID") AND url_imagem_referencia = "" then
								        capa_audio = true
    							    end if
                                Next
                            Next
						end if
					Next
				end if
				if url_imagem_referencia = "" then
					For Each xmlCampos In xmlFicha.childNodes
						if xmlCampos.nodeName = "SITES" AND url_imagem_referencia = "" then
                            For Each xmlSubCampos In xmlCampos.childNodes
                                For Each xmlItem In xmlSubCampos.childNodes
                                    if url_imagem_referencia = "" then
							            url_site = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
							            url_imagem_referencia = url_site
                                        codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
                                        dspace = xmlItem.attributes.getNamedItem("DSpace").value
                                    end if
                                Next
                            Next
						end if
					Next
				end if
			end if
		Next
	end if

	titulo_autor_sem_capa = titulo_autor_sem_capa & titulo_f_sem_capa & desc_autor_sem_capa & "</div>"

	if (url_imagem_referencia <> "") then
		titulo_autor_sem_capa = titulo_autor_sem_capa & "<div class='mascara_capa_fantasia_transparente' data-url=""" & url_imagem_referencia & """ onclick=""abrirLink(this);ContarAcesso("&codigo_obra&","&codMidia&",1,"&dspace&");""></div>"
	end if
	'*************************************
	'FIM EXTRAÇÃO DE TÍTULO E AUTOR, QUANDO NÃO EXISTIR CAPA PARA O ELEMENTO PESQUISADO
	'*************************************

	'*************************************
	'Miniatura
	'*************************************
	if (global_exibe_capa = 1) then
		if(tem_capa = "1") then
			Response.Write "<table class='max_width'>"
			Response.Write "<td style='width: 110px;' class='td_left_top'>"
			if url_imagem_referencia <> "" then
				Response.Write "<a href='" & url_imagem_referencia & "' target='_blank' onclick='ContarAcesso("&codigo_obra&","&codMidia&",1,"&dspace&");'>"
			end if
			Response.Write "<img alt='' src='asp/capa.asp?obra="&codigo_obra&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"'>"
			if url_imagem_referencia <> "" then
				Response.Write "</a>"
			end if
			Response.Write "</td>"
			Response.Write "<td style='vertical-align: top;'>"
		else
			if (codigo_material = 46 or codigo_material = 35) then
				Response.Write "<table class='max_width'>"
				Response.Write "<td style='width: 110px;' class='td_left_top'>"
				if url_imagem_referencia <> "" then
					Response.Write "<a href='" & url_imagem_referencia & "' target='_blank' onclick='ContarAcesso("&codigo_obra&","&codMidia&",1,"&dspace&");'>"
				end if
				Response.Write "<div class='capa_audio icon_capa'></div>" 
				if url_imagem_referencia <> "" then
					Response.Write "</a>"
				end if
				Response.Write "</td>"
				Response.Write "<td style='vertical-align: top;'>"
			else
				Response.Write "<table class='max_width'>"
				Response.Write "<td style='width: 110px;' class='td_left_top'>"
				Response.Write titulo_autor_sem_capa
				Response.Write "<div class='capa_fantasia icon_capa'></div>"
				Response.Write "</td>"
				Response.Write "<td style='vertical-align: top;'>"
			end if
		end if
	end if

	'************************************************************
	' MONTA O DETALHE PARA OBRAS OU PERIÓDICOS
	'************************************************************
	if cInt(Tipo) <> 2 AND cInt(Tipo) >= 0 then        
		if xmlRoot.nodeName = "FICHA_RESUMIDA" then
			For Each xmlPNome In xmlRoot.childNodes
				'************************************************************
				' FICHA SOPHIA
				'************************************************************
				if xmlPNome.nodeName = "FICHA" then
					Response.Write "<table class='max_width table-ficha-detalhes'>"
					class_detalhe = "td_detalhe_valor"
					For Each xmlCampos In xmlPNome.childNodes
						'************************************************************
						' INFORMAÇÃO DA PUBLICAÇÃO
						'************************************************************
						if xmlCampos.nodeName = "INF_PUBLICACAO" then
							inf_publicacao = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_inf_publicacao = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_inf_publicacao&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(inf_publicacao,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' IDIOMA (BNBD)
						'************************************************************
						if xmlCampos.nodeName = "IDIOMA" then
							idioma = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_idioma = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_idioma&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(idioma,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' ISBN
						'************************************************************
						if xmlCampos.nodeName = "ISBN" then
							isbn = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_isbn = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao' >"&desc_isbn&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(isbn,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' ISSN
						'************************************************************
						if xmlCampos.nodeName = "ISSN" then
							issn = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_issn = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_issn&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(issn,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' CDD (BNBD)
						'************************************************************
						if xmlCampos.nodeName = "CDD" then
							cdd = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_cdd = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_cdd&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(cdd,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' OBJETO DIGITAL (BNBD)
						'************************************************************
						if xmlCampos.nodeName = "OBJETO_DIGITAL" then
							objeto_dig = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_objeto_dig = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_objeto_dig&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(objeto_dig,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
                        '************************************************************
						' NUMERO CHAMADA LOCAL 
						'************************************************************
						if xmlCampos.nodeName = "NUMERO_CHAMADA_LOCAL" then
							numero_chamada_local = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_numero_chamada_local = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_numero_chamada_local&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(numero_chamada_local,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' NUMERO CHAMADA
						'************************************************************
						if xmlCampos.nodeName = "CHAMADA" then
							desc_chamada = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_chamada&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
							Response.Write "</tr>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' CLASSIFICACAO, NOTACAO, COMPLEMENTO, OUTRAS
								'************************************************************
								desc_chamada = xmlSubCampos.attributes.getNamedItem("Descricao").value
								chamada = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								Response.Write "<tr>"
								Response.Write "<td class='td_detalhe_descricao_imagem'>"
								Response.Write "<div class='joinbottom icon_20'></div>" 
								Response.Write "</td>"
								Response.Write "<td class='td_detalhe_descricao_recuo'>"
								Response.Write desc_chamada&"&nbsp;"
								Response.Write "</td>"
								Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & replace(chamada,chr(10),"<br/>") & "&nbsp;</div></td>"
								Response.Write "</tr>"
							Next
						end if
						'************************************************************
						' LOC. ORIGINAL (BNBD)
						'************************************************************
						if xmlCampos.nodeName = "LOC_ORIGINAL" then
							loc_original = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_loc_original = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_loc_original&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(loc_original,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'*************************************
						'LOCALIZAÇÃO (BN)
						'*************************************
						if xmlCampos.nodeName = "LOCALIZACAO" then
							loc = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_loc = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_loc&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(loc,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' ENTRADA PRINCIPAL
						'************************************************************
						if xmlCampos.nodeName = "ENT_PRINC" then
							codigo_ent_princ = trim(xmlCampos.attributes.getNamedItem("Codigo").value)
							tipo_ent_princ = trim(xmlCampos.attributes.getNamedItem("Tipo").value)
							desc_ent_princ = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							valor_ent_princ = TrocaTagMarcador(trim(xmlCampos.attributes.getNamedItem("Valor").value))
							busca_ent_princ = RemoveTagMarcador(trim(xmlCampos.attributes.getNamedItem("Valor").value))
							desc_princ_ent_princ = replace(replace(replace(busca_ent_princ," ","_"),"&#39;","_#39;"),"'","\'")
							if (somente_detalhe = false) then
								nome_funcao = "LinkBuscaAutor"
								hint = getTermo(global_idioma, 1558, "Buscar todos os registros deste autor", 0)&"..."
								if xmlCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
									autInfo = "<td class='autImg'><div class='autImg' id='autImg' onClick=LinkAutInfo('detalhes',"&codigo_ent_princ&",'"&desc_princ_ent_princ&"',"&tipo_ent_princ&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1571, "Mostrar informações sobre este autor", 0)&"...'>&nbsp;</div></td>"
								else
									autInfo = ""
								end if
                                identificadoresInfo = ""
                                For Each xmlIdentificadores in xmlCampos.childNodes
                                    %><!--#include file ="identificadoresDetalhe.asp"--><%                                                                                                     
                                next
                                autInfo = autInfo & identificadoresInfo
								temp = "<a class='link_custom' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_ent_princ&",'"&desc_princ_ent_princ&"',"&iIndexSrv&");"">"&replace(valor_ent_princ,chr(10),"<br/>") &"</a>"
								autor = "<table class='remover_bordas_padding autLink'><tr><td id='autLinkTab' class='td_left_middle'>"&temp&"&nbsp;</td>"&autInfo&"</tr></table>"
							else
								autor = valor_ent_princ
							end if
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_ent_princ&"</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & autor & "</td>"
							Response.Write "</tr>"
						end if
						'*************************************
						'AUTORIA (ENTRADA PRINCIPAL e ENTRADA(S) SECUNDÁRIA(S))
						'Adequação para o cliente IMS
						'*************************************
						if xmlCampos.nodeName = "AUTORIA" then
							
							autor = ""
							desc_ent_princ = replace(replace(replace(trim(xmlCampos.attributes.getNamedItem("Desc_Autoria").value)," ","_"),"&#39;","_#39;"),"'","\'")
							sequencialOld = sequencial
							For Each xmlSubCampos In xmlCampos.childNodes
								sequencial = sequencial + 1
								desc_autor = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								autor_busca = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Desc_Princ").value)
								autor_tipo = xmlSubCampos.attributes.getNamedItem("Tipo").value
								autor_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							

								nome_funcao = "LinkBuscaAutor"
								
								autor_formatado = replace(replace(replace(replace(autor_busca," ","_"),"<",""),">",""),"'","\'")
								
								if xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
									autInfo = "<td class='autImg'><div class='autImg' id='autImg"&sequencial&"' onClick=LinkAutInfo('linkAutInfo',"&autor_codigo&",'"&autor_formatado&"',"&autor_tipo&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1571, "Mostrar informações sobre este autor", 0)&"...'>&nbsp;</div></td>"
								else
									autInfo = ""
								end if	                                
								autor = autor & "<table class='remover_bordas_padding autLink'><tr><td id='autLinkTab"&sequencial&"' class='td_left_middle'><a class='link_classic2' title='"&getTermo(global_idioma, 1558, "Buscar todos os registros deste autor", 0)&"...' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&autor_codigo&",'"&autor_formatado&"',"&iIndexSrv&");"">"&desc_autor&"</a>&nbsp;</td>"&autInfo&"</tr></table>"
							Next
							sequencial = sequencialOld
							
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_ent_princ&"</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & autor & "</td>"
							Response.Write "</tr>"
                            
						end if
						'************************************************************
						' TITULO
						'************************************************************
						if xmlCampos.nodeName = "TITULO" then
							titulo_f = Replace(Replace(xmlCampos.attributes.getNamedItem("Valor").value, "<", "&lt;"), ">", "&gt;")							
							titulo = TrocaTagMarcador(titulo_f)
							desc_titulo = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_titulo&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'><b>" & replace(titulo,chr(10),"<br/>") & "&nbsp;</b></td>"
							Response.Write "</tr>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' OUTROS TITULOS
								'*********************************************************			***
								desc_titulo = xmlSubCampos.attributes.getNamedItem("Descricao").value
								titulo = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								titulo = replace(titulo,chr(10),"<br/>")
								if (Not xmlSubCampos.attributes.getNamedItem("Codigo") Is Nothing) then
									tipo = "1"
									if (Not xmlSubCampos.attributes.getNamedItem("Tipo") Is Nothing) then
										tipo = xmlSubCampos.attributes.getNamedItem("Tipo").value
								end if
									titulo = "<a href='javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1," & trim(xmlSubCampos.attributes.getNamedItem("Codigo").value) & ",1,""link_detalhe""," & tipo & ");'>" & titulo & "</a>"
								end if
								Response.Write "<tr>"
								Response.Write "<td class='td_detalhe_descricao_imagem'>"
								Response.Write "<div class='joinbottom icon_20'></div>" 
								Response.Write "</td>"
								Response.Write "<td class='td_detalhe_descricao_recuo'>"
								Response.Write desc_titulo&"&nbsp;"
								Response.Write "</td>"
								Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & titulo & "&nbsp;</div></td>"
								Response.Write "</tr>"
							Next
						end if
						'************************************************************
						' TÍTULO UNIFORME
						'************************************************************
						if xmlCampos.nodeName = "TITULO_UNI" then
							titulo = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_titulo = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_titulo&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(titulo,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' TÍTULO ABREVIADO
						'************************************************************
						if xmlCampos.nodeName = "TITULO_ABREV" then
							titulo = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_titulo = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_titulo&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(titulo,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
                        '************************************************************
						' Classificação DEWEY
						'************************************************************
                        if (global_numero_serie = 5592) then
                            if xmlCampos.nodeName = "DEWEY" then
							    classificacao_dewey = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							    desc_classificacao_dewey = xmlCampos.attributes.getNamedItem("Descricao").value
							    Response.Write "<tr>"
							    Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_classificacao_dewey&"&nbsp;</td>"
							    Response.Write "<td class='"&class_detalhe&"'>" & replace(classificacao_dewey,chr(10),"<br/>") & "&nbsp;</td>"
							    Response.Write "</tr>"
                               
							    For Each xmlSubCampos In xmlCampos.childNodes
                                    '************************************************************
						            ' DEWEY - EDIÇÃO
						            '************************************************************	
                                    desc_edicao = xmlSubCampos.attributes.getNamedItem("Descricao").value
								    edicao = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								    Response.Write "<tr>"
								    Response.Write "<td class='td_detalhe_descricao_imagem'>"
								    Response.Write "<div class='joinbottom icon_20'></div>" 
								    Response.Write "</td>"
								    Response.Write "<td class='td_detalhe_descricao_recuo'>"
								    Response.Write desc_edicao&"&nbsp;"
								    Response.Write "</td>"
								    Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & replace(edicao,chr(10),"<br/>") & "&nbsp;</div></td>"
								    Response.Write "</tr>"
                                Next
						    end if
                        end if

                        '************************************************************
						' EDIÇÃO
						'************************************************************
						if xmlCampos.nodeName = "EDICAO" then
							edicao = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_edicao = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_edicao&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(edicao,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
                        '************************************************************
						' CARTOGRAFIA
						'************************************************************
						if xmlCampos.nodeName = "CARTOGRAFIA" then
							cartografia = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_catografia = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_catografia&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(cartografia,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' IMPRENTA
						'************************************************************
						if xmlCampos.nodeName = "IMPRENTA" then
							imprenta = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_imprenta = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_imprenta&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(imprenta,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' DESCRIÇÃO FISICA
						'************************************************************
						if xmlCampos.nodeName = "DESC_FISICA" then
							desc_fisica = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_desc_fisica = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_desc_fisica&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(desc_fisica,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' FORMA DO REGISTRO
						'************************************************************
						if xmlCampos.nodeName = "FORMA_REGISTRO" then
							desc_fisica = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_desc_fisica = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='esquerda td_detalhe_descricao'>"&desc_desc_fisica&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(desc_fisica,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' SÉRIE
						'************************************************************
						if xmlCampos.nodeName = "SERIE" then
							serie = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_serie = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_serie&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(serie,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' PERIODICIDADE
						'************************************************************
						if xmlCampos.nodeName = "PERIODICIDADE" then
							periodicidade = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_periodicidade = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_periodicidade&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(periodicidade,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' PERIODICIDADE ANTERIOR
						'************************************************************
						if xmlCampos.nodeName = "PERIODICIDADE_ANT" then
							periodicidade_ant = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_periodicidade_ant = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_periodicidade_ant&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(periodicidade_ant,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' NOTAS
						'************************************************************
						if xmlCampos.nodeName = "NOTAS" then
							desc_nota = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_nota&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
							Response.Write "</tr>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' CARREGA NOTAS EXISTENTES
								'************************************************************
								desc_nota = xmlSubCampos.attributes.getNamedItem("Descricao").value
								nota = xmlSubCampos.attributes.getNamedItem("Valor").value
								posUrlIni = InStr(nota, "#HREF#")
								if (posUrlIni > 0) then
									posUrlFinal = InStr(nota, "#/HREF#") + 5
									nota = Mid(nota, 1, (posUrlIni -1)) + RemoveTagMarcador(Mid(nota,posUrlIni,(posUrlFinal-posUrlIni+1))) + Mid(nota,(posUrlFinal + 1),(Len(nota)-posUrlFinal))
								end if
								nota = TrocaTagMarcador(nota)
								nota = replace(replace(replace(replace(nota,"#URL#","<a class='link_classic2' target='_blank' "), "#HREF#", "href='"), "#/HREF#", "'>"), "#/URL#", "</a>")
								Response.Write "<tr>"
								Response.Write "<td class='td_detalhe_descricao_imagem'>"
								Response.Write "<div class='joinbottom icon_20'></div>" 
								Response.Write "</td>"
								Response.Write "<td class='td_detalhe_descricao_recuo'>"
								Response.Write desc_nota&"&nbsp;"
								Response.Write "</td>"
								Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & TrocaTagEspaco(nota) & "&nbsp;</div></td>"
								Response.Write "</tr>"
							Next
						end if
						'************************************************************
						' ASSUNTOS
						'************************************************************
						if xmlCampos.nodeName = "ASSUNTOS" then
							seq = 1
							assunto = ""
							desc_assunto = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
								codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
								valor_assunto = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								busca_assunto = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								desc_princ_assunto = replace(replace(replace(busca_assunto," ","_"),"&#39;","_#39;"),"'","\'")
								seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)
								if (somente_detalhe = false) then
									nome_funcao = "LinkBuscaAssunto"
									hint = getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"..."
									if xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assInfo = "<span class='assuntoImg span_imagem icon_16' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</span>"
    									else
										    assInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
										end if                                        
									else
										assInfo = ""
									end if
									temp = seq_assunto&" <a class='link_custom' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_assunto&",'"&desc_princ_assunto&"',"&iIndexSrv&");"">"&replace(valor_assunto,chr(10),"<br/>") &"</a>"
								else
									assInfo = ""
									temp = seq_assunto & " " & valor_assunto
								end if   
                                identificadoresInfo = ""                                
                                For Each xmlIdentificadores in xmlSubCampos.childNodes                                    
                                    %><!--#include file ="identificadoresDetalhe.asp"--><%                                                                                                      
                                next                                
                                assInfo = assInfo & identificadoresInfo
								if assunto <> "" then
									'Adequação ABL - Exibir itens um em frente do outro
									if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
										assunto = assunto & "&nbsp;" & temp & "&nbsp;" & assInfo
									else
										assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0;'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
									end if
								else
									'Adequação ABL - Exibir itens um em frente do outro
									if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
										assunto = assunto & temp & "&nbsp;" & assInfo
									else
										assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0;'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
									end if
								end if
								seq = seq + 1
							Next
							Response.Write "<td class='"&class_detalhe&"'>" & assunto & "</td>"
							Response.Write "</tr>"
                            
						end if
						'************************************************************
						' ENTRADA SECUNDÁRIA
						'************************************************************
						if xmlCampos.nodeName = "ENTRADA_SECUNDARIA" then
							seq = 1
							autor = ""
							desc_ent_sec = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_ent_sec&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								tipo_ent_sec = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
								codigo_ent_sec = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
								valor_ent_sec = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								seq_ent_sec = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)
                                serieAutoria = trim(xmlSubCampos.attributes.getNamedItem("SERIE").value)

								if (somente_detalhe = false) then
									nome_funcao = "LinkBuscaAutor"
                                    if (serieAutoria = 1) then
                                        hint = getTermo(global_idioma, 8363, "Buscar todas obras desta entrada secundária", 0)&"..."
                                        hintInfo = getTermo(global_idioma, 8364, "Mostrar informações sobre esta entrada secundária", 0)&"..."
								        busca_ent_sec = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Desc_Princ_Serie").value))
								        desc_princ_ent_sec = replace(replace(replace(busca_ent_sec," ","_"),"&#39;","_#39;"),"'","\'")
                                    else 
									    hint = getTermo(global_idioma, 1558, "Buscar todos os registros deste autor", 0)&"..."
                                        hintInfo = getTermo(global_idioma, 1571, "Mostrar informações sobre este autor", 0)&"..."
								        busca_ent_sec = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Desc_Princ").value))
								        desc_princ_ent_sec = replace(replace(replace(busca_ent_sec," ","_"),"&#39;","_#39;"),"'","\'")
                                    end if
									if xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
										'Adequação ABL - Exibir itens um em frente do outro
										if (global_numero_serie = 2372) or (global_numero_serie = 2635) then
											autInfo = "<span class='autImg span_imagem icon_16' id='autImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_ent_sec&",'"&desc_princ_ent_sec&"',"&tipo_ent_sec&") style='cursor:pointer;' title='"&hintInfo&"'>&nbsp;</span>"
										else
										    autInfo = "<td class='autImg'><div class='autImg' id='autImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_ent_sec&",'"&desc_princ_ent_sec&"',"&tipo_ent_sec&") style='cursor:pointer;' title='"&hintInfo&"'>&nbsp;</div></td>"
										end if
									else
										autInfo = ""
									end if
                                    
									temp = seq_ent_sec&" <a class='link_custom' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_ent_sec&",'"&desc_princ_ent_sec&"',"&iIndexSrv&","&serieAutoria&");"">"&replace(valor_ent_sec,chr(10),"<br/>") &"</a>"
								else
									autInfo = ""
									temp = seq_ent_sec & " " & valor_ent_sec
								end if 
                                identificadoresInfo = ""                                
                                For Each xmlIdentificadores in xmlSubCampos.childNodes                                       
                                    %><!--#include file ="identificadoresDetalhe.asp"--><%  
                                next                                
                                autInfo = autInfo & identificadoresInfo	
								if autor <> "" then						
									'Adequação ABL - Exibir itens um em frente do outro
									if (global_numero_serie = 2372) or (global_numero_serie = 2635) then
										autor = autor & "&nbsp;" & temp & "&nbsp;" & autInfo
									else
										autor = autor & "<table class='autLink' style='border-spacing: 1px; padding: 0;'><tr><td id='autLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & autInfo & "</tr></table>"
									end if
								else
									'Adequação ABL - Exibir itens um em frente do outro
									if (global_numero_serie = 2372) or (global_numero_serie = 2635) then
										autor = temp & "&nbsp;" & autInfo
									else
									autor = "<table class='autLink' style='border-spacing: 1px; padding: 0;'><tr><td id='autLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & autInfo & "</tr></table>"
									end if
								end if
								seq = seq + 1
							Next
							Response.Write "<td class='"&class_detalhe&"'>" & autor & "</td>"
							Response.Write "</tr>"
						end if
                        '************************************************************
						' SubLocalização (BNBD)
						'************************************************************
						if xmlCampos.nodeName = "SUBLOCALIZACAO" then
							sublocalizacao = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_sublocalizacao = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_sublocalizacao&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(sublocalizacao,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
                        '************************************************************
						' TITULOS_NAO_CONTROLADOS
						'************************************************************
                        if xmlCampos.nodeName = "TITULOS_NAO_CONTROLADOS" then
                                titulo = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							    desc_titulo = xmlCampos.attributes.getNamedItem("Descricao").value
							    Response.Write "<tr>"
							    Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_titulo&"&nbsp;</td>"
							    Response.Write "<td class='"&class_detalhe&"'>" & replace(titulo,chr(10),"<br/>") & "&nbsp;</td>"
							    Response.Write "</tr>"
                        end if
						'************************************************************
						' PUBLICADO COM
						'************************************************************
						if xmlCampos.nodeName = "PUBLICADO_COM" then
							desc_publicado = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_publicado&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
							Response.Write "</tr>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' PUBLICADO
								'************************************************************
								desc_publicado = xmlSubCampos.attributes.getNamedItem("Descricao").value
								publicado = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								Response.Write "<tr>"
								Response.Write "<td class='td_detalhe_descricao_imagem'>"
								Response.Write "<div class='joinbottom icon_20'></div>" 
								Response.Write "</td>"
								Response.Write "<td class='td_detalhe_descricao_recuo'>"
								Response.Write desc_publicado&"&nbsp;"
								Response.Write "</td>"
								Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & replace(publicado,chr(10),"<br/>") & "&nbsp;</div></td>"
								Response.Write "</tr>"
							Next
						end if
						'************************************************************
						' RESUMO
						'************************************************************
						if xmlCampos.nodeName = "RESUMO" then
							nome_funcao = "LinkBuscaResumo"
							hint = getTermo(global_idioma, 1565, "Exibir resumo completo", 0)&"..."
							modo = "res"
							resumo = xmlCampos.attributes.getNamedItem("Valor").value
							desc_resumo = xmlCampos.attributes.getNamedItem("Descricao").value
							Mais = trim(xmlCampos.attributes.getNamedItem("Mais").value)
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_resumo&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & replace(resumo,chr(10),"<br/>")
							if cStr(Mais) = "1" then
								Response.Write "&nbsp;...</div><div class='direita'><span class='span_imagem div_imagem_right_3 icon_9 mais-png '></span>"
                                Response.Write "<a class='link_classic2' title='"&hint&"'onClick="&nome_funcao&"("&codigo&",'"&modo&"',"&iIndexSrv&") href=""javascript:"&nome_funcao&"("&codigo&",'"&modo&"');"">Ler mais</a>&nbsp;</div></td>"
							else
								Response.Write "</div></td>"
							end if
							Response.Write "</tr>"
						end if
						if xmlCampos.nodeName = "NOTAS_FONTE" then
							desc_nota = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_nota&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
							Response.Write "</tr>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' CARREGA NOTAS EXISTENTES
								'************************************************************
								desc_nota = xmlSubCampos.attributes.getNamedItem("Descricao").value
								nota = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								Response.Write "<tr>"
								Response.Write "<td class='td_detalhe_descricao_imagem'>"
								Response.Write "<div class='joinbottom icon_20'></div>" 
								Response.Write "</td>"
								Response.Write "<td class='td_detalhe_descricao_recuo'>"
								Response.Write desc_nota&"&nbsp;"
								Response.Write "</td>"
								Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & replace(nota,chr(10),"<br/>") & "&nbsp;</div></td>"
								Response.Write "</tr>"
							Next
						end if
						'************************************************************
						' ACADEMICO
						'************************************************************
						if ((global_versao = vSOPHIA) and (global_academico = 1) and (xmlCampos.nodeName = "ACADEMICO")) then
							academico = ""
							For Each xmlSubCampos In xmlCampos.childNodes
								academico = academico & xmlSubCampos.attributes.getNamedItem("Valor").value & "<br/>"
							Next
							
							if (trim(academico) <> "") then
								Response.Write "<tr>"
								Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&xmlCampos.attributes.getNamedItem("Descricao").value&"&nbsp;</td>"
								Response.Write "<td class='"&class_detalhe&"'>" & academico &"</td>"
								Response.Write "</tr>"
							end if
						end if
						'************************************************************
						' SITES RELACIONADOS
						'************************************************************
						if xmlCampos.nodeName = "SITES" then
							desc_sites = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_sites&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
							Response.Write "</tr>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' SITES
								'************************************************************
								desc_site = xmlSubCampos.attributes.getNamedItem("Descricao").value
								site_c = ""
								For Each xmlItem In xmlSubCampos.childNodes
									codMidia = xmlItem.attributes.getNamedItem("Codigo").value
									plataforma = xmlItem.attributes.getNamedItem("Plataforma").value
									site = TrocaTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
									site_url = RemoveTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
                                    dspace = xmlItem.attributes.getNamedItem("DSpace").value
									if site_c <> "" then
										site_c = site_c & "<br/>"
									end if
									if (plataforma <> "") then
										site_c = site_c & "<a class='link_classic2' href='javascript:abreMidiaEspecifica(" & Tipo & "," & codigo_obra & "," & codMidia & "," & iIndexSrv & ")'>" & site & "</a>"
									else
										site_c = site_c & "<a class='link_classic2' href='" & site_url & "' target='_blank' onclick='ContarAcesso("&codigo_obra&","&codMidia&",1,"&dspace&");'>" & site & "</a>"
									end if
								Next	
								
								Response.Write "<tr>"
								Response.Write "<td class='td_detalhe_descricao_imagem'>"
								Response.Write "<div class='joinbottom icon_20'></div>" 
								Response.Write "</td>"
								Response.Write "<td class='td_detalhe_descricao_recuo'>"
								Response.Write desc_site&"&nbsp;"
								Response.Write "</td>"
								Response.Write "<td class='"&class_detalhe&"'>" & site_c & "&nbsp;</td>"
								Response.Write "</tr>"
							Next
						end if
						'************************************************************
						' URL
						'************************************************************
						if xmlCampos.nodeName = "URL" then
							if (repositorio_institucional = 1) then
								url = global_cfg_url_rep_institucional & "/"
							else
								url = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							end if
							desc_url = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_url&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & url & "index.asp?codigo_sophia=" & codigo_obra & "</td>"
							Response.Write "</tr>"
						end if									
						'************************************************************
						' FONTES ( apenas para analíticas )
						'************************************************************
						if xmlCampos.nodeName = "FONTES" then
							For Each xmlFontes In xmlCampos.childNodes
								if xmlFontes.nodeName = "FONTE_OBRA" then
									desc_fonte = xmlFontes.attributes.getNamedItem("DESCRICAO").value
									Response.Write "<tr>"
									Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_fonte&"&nbsp;</td>"
									Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
									Response.Write "</tr>"

									For Each xmlSubCampo In xmlFontes.childNodes
										sDesc_fonte    = xmlSubCampo.attributes.getNamedItem("DESCRICAO").value
										sTit_fonte     = TrocaTagMarcador(xmlSubCampo.attributes.getNamedItem("TITULO").value)
										sComp_fonte    = TrocaTagMarcador(xmlSubCampo.attributes.getNamedItem("COMPLEMENTO").value)
										sPreComp_fonte = TrocaTagMarcador(xmlSubCampo.attributes.getNamedItem("PRE_COMP").value)
										sTipo_fonte    = xmlSubCampo.attributes.getNamedItem("TIPO").value							
										sCod_fonte     = xmlSubCampo.attributes.getNamedItem("CODIGO").value							
															
										sTit_fonte = "<a class='link_classic2' title='"&getTermo(global_idioma, 1567, "Ver detalhes da obra", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,0,"&sCod_fonte&","&pagina&",'periodico',"&sTipo_fonte&");"">"&sTit_fonte&"</a>"
		
										Response.Write "<tr>"
										Response.Write "<td class='td_detalhe_descricao_imagem'>"
										Response.Write "<div class='joinbottom icon_20'></div>" 
										Response.Write "</td>"
										Response.Write "<td class='td_detalhe_descricao_recuo'>"
										Response.Write sDesc_fonte&"&nbsp;"
										Response.Write "</td>"
										Response.Write "<td class='"&class_detalhe&"'>" & sPreComp_fonte&sTit_fonte&sComp_fonte & "&nbsp;</td>"
										Response.Write "</tr>"
									Next
								elseif xmlFontes.nodeName = "FONTE_PER" then
									desc_fonte = xmlFontes.attributes.getNamedItem("DESCRICAO").value
									
									Response.Write "<tr>"
									Response.Write "<td colspan='2' class='td_detalhe_descricao'>"&desc_fonte&"&nbsp;</td>"
									Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
									Response.Write "</tr>"

									For Each xmlSubCampo In xmlFontes.childNodes
										sDesc_fonte = xmlSubCampo.attributes.getNamedItem("DESCRICAO").value
										sTit_fonte  = TrocaTagMarcador(xmlSubCampo.attributes.getNamedItem("TITULO").value)
										sComp_fonte = TrocaTagMarcador(xmlSubCampo.attributes.getNamedItem("COMPLEMENTO").value)
										sTipo_fonte = xmlSubCampo.attributes.getNamedItem("TIPO").value							
										sCod_fonte  = xmlSubCampo.attributes.getNamedItem("CODIGO").value							
															
										if (sTipo_fonte = "0") then
											sCodEx	   = xmlSubCampo.attributes.getNamedItem("EXEMPLAR").value
											sTit_fonte = "<a class='link_classic2' title='"&getTermo(global_idioma, 1568, "Ver detalhes do periódico", 0)&"...' href=""javascript:LinkDetalhesPeriodico(parent.hiddenFrame.modo_busca,1,0,"&sCod_fonte&",1,'periodico',"&sCodEx&",0);"">"&sTit_fonte&"</a>"
										else
											sTit_fonte = "<a class='link_classic2' title='"&getTermo(global_idioma, 1568, "Ver detalhes do periódico", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,0,"&sCod_fonte&",1,'periodico',"&sTipo_fonte&");"">"&sTit_fonte&"</a>"
										end if
		
										Response.Write "<tr>"
										Response.Write "<td class='td_detalhe_descricao_imagem'>"
										Response.Write "<div class='joinbottom icon_20'></div>" 
										Response.Write "</td>"
										Response.Write "<td class='td_detalhe_descricao_recuo'>"
										Response.Write sDesc_fonte&"&nbsp;"
										Response.Write "</td>"
										Response.Write "<td class='"&class_detalhe&"'>" & sTit_fonte&sComp_fonte & "&nbsp;</td>"
										Response.Write "</tr>"
									Next
								End if
							Next								
						end if
					Next

					'Para resolver o problema no IE, quando utilizado colspan é ignorado o width da coluna
					Response.Write "<tr><td style='width: 19px;'></td><td style='width: 141px'></td><td></td></tr>"
					Response.Write "</table>"
					if global_exibe_capa = 1 then
						Response.Write "</td></tr>"
						Response.Write "</table>"
					end if
				'************************************************************
				' TESES E DISSERTAÇÕES
				'************************************************************
				elseif xmlPNome.nodeName = "TESE" then
					if (global_exibe_capa = 1) then
						style = " padding-left: 114px"
					else
						style = ""
					end if
					Response.Write "</td></tr>"
					Response.Write "<tr><td class='centro' style='padding-bottom: 6px;" & style & "'><br/><b>"&getTermo(global_idioma, 135, "Outras informações", 0)&"</b></td></tr>"
					Response.Write "<tr><td>"
					Response.Write "<table class='max_width' style='border-spacing: 1px; padding: 0;" & style & "'>"
					For Each xmlCampos In xmlPNome.childNodes
						'************************************************************
						' Instituição de Defesa
						'************************************************************
						if xmlCampos.nodeName = "INST_DEFESA" then
							inst_defesa = xmlCampos.attributes.getNamedItem("Valor").value
							desc_inst_defesa = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td class='td_detalhe_descricao'>"&desc_inst_defesa&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & inst_defesa & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' Programa
						'************************************************************
						if xmlCampos.nodeName = "PROGRAMA" then
							programa = xmlCampos.attributes.getNamedItem("Valor").value
							desc_programa = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td class='td_detalhe_descricao' >"&desc_programa&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & programa & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' Área
						'************************************************************
						if xmlCampos.nodeName = "AREA" then
							area = xmlCampos.attributes.getNamedItem("Valor").value
							desc_area = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td class='td_detalhe_descricao' >"&desc_area&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & area & "&nbsp;</td>"
							Response.Write "</tr>"
						end if

						'************************************************************
						' Titulação
						'************************************************************
						if xmlCampos.nodeName = "TITULACAO" then
							titulacao = xmlCampos.attributes.getNamedItem("Valor").value
							desc_titulacao = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td class='td_detalhe_descricao'>"&desc_titulacao&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & titulacao & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' Data de Defesa
						'************************************************************
						if xmlCampos.nodeName = "DATA_DEFESA" then
							dtdefesa = xmlCampos.attributes.getNamedItem("Valor").value
							desc_dtdefesa = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td class='td_detalhe_descricao'>"&desc_dtdefesa&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & dtdefesa & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
                        '************************************************************
						' Adequação ITA - Cursos dos autores
						'************************************************************
                        if xmlCampos.nodeName = "CURSOS_AUTORES" then
							seq = 1
							curso = ""
                            desc_cursos = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
                            Response.Write "<td class='td_detalhe_descricao'>"&desc_cursos&"&nbsp;</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								valor_curso = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								seq_curso = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)
								temp = seq_curso & ". " & valor_curso
                                curso = curso & "<table class='autLink' style='border-spacing: 1px; padding: 0;'><tr><td id='autLinkTab"&seq&"' class='td_left_middle'>" & temp & "</td></tr></table>"
                                
								seq = seq + 1
							Next

							Response.Write "<td class='"&class_detalhe&"'>" & curso & "</td>"
							Response.Write "</tr>"
						end if

					Next
					Response.Write "</table>"
				'************************************************************
				' CAMPOS OPCIONAIS
				'************************************************************
				elseif xmlPNome.nodeName = "DGM" then
					TEM_DGM = 0
					if (global_exibe_capa = 1) then
						style = " padding-left: 114px"
					else
						style = ""
					end if
					For Each xmlCampos In xmlPNome.childNodes
						if xmlCampos.nodeName = "CAMPO" then
							if TEM_DGM = 0 then
								Response.Write "</td></tr>"				
								Response.Write "<tr><td class='centro' style='padding-bottom: 6px;" & style & "'><br/><b>"&getTermo(global_idioma, 1570, "Descrição complementar do material", 0)&"</b></td></tr>"
								Response.Write "<tr><td>"
								Response.Write "<table class='max_width' style='border-spacing: 1px; padding: 0;" & style & "'>"
								TEM_DGM = 1
							end if
							dgm = xmlCampos.attributes.getNamedItem("Valor").value
							desc_dgm = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td class='td_detalhe_descricao'>"&desc_dgm&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & replace(dgm,chr(10),"<br/>") & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
					Next
					if TEM_DGM = 1 then
						Response.Write "</table>"
					end if
				'************************************************************
				' SERVIÇOS
				'************************************************************
				elseif xmlPNome.nodeName = "SERVICOS" then
					For Each xmlServicos In xmlPNome.childNodes
						'*************************************
						'SERVIÇOS HABILITADOS PARA A OBRA
						'*************************************
						if xmlServicos.nodeName = "LINK_SELECIONAR" AND (somente_detalhe = false) and (global_esconde_menu = 0) then
							serv_selecao = true
						end if
						if xmlServicos.nodeName = "LINK_RESERVAR" and (global_esconde_menu = 0) then
							serv_reserva = (repositorio_institucional = 0)
						end if
						if xmlServicos.nodeName = "LINK_AQUISICAO" AND (somente_detalhe = false)  and (global_esconde_menu = 0) then
							serv_aquisicao = (repositorio_institucional = 0)
						end if
						if xmlServicos.nodeName = "LINK_MIDIAS" then
							serv_midias = true
							serv_midias_desc = xmlServicos.attributes.getNamedItem("Descricao").value
							serv_midias_mus = xmlServicos.attributes.getNamedItem("Audio").value
						end if
						if xmlServicos.nodeName = "LINK_REF_BIB" then
							serv_refbib = true
						end if
						if xmlServicos.nodeName = "LINK_BIB_COMP" then
							serv_bibcomp = true
						end if
						if xmlServicos.nodeName = "LINK_ANALITICA" then
							serv_analitica = true
							serv_analitica_desc = xmlServicos.attributes.getNamedItem("Descricao").value
						end if
					Next
				end if
			Next
		end if
	'************************************************************
	' MONTA O DETALHE PARA LEGISLAÇÃO
	'************************************************************
	elseif Tipo = 2 then
		if xmlRoot.nodeName = "FICHA_RESUMIDA" then
			For Each xmlPNome In xmlRoot.childNodes
				'************************************************************
				' FICHA LEGISLAÇÃO
				'************************************************************
				if xmlPNome.nodeName = "FICHA" then
					Response.Write "<table class='max_width' style='border-spacing: 1px; padding: 0;'>"
					class_detalhe = "td_detalhe_valor"
					For Each xmlCampos In xmlPNome.childNodes
						'************************************************************
						' NORMA
						'************************************************************
						if xmlCampos.nodeName = "NORMA" then
							desc_norma = xmlCampos.attributes.getNamedItem("Descricao").value
							norma = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>" & desc_norma & "&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'><b>" & norma & "&nbsp;</b></td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' APELIDO
						'************************************************************
						if xmlCampos.nodeName = "APELIDO" then
							desc_apelido = xmlCampos.attributes.getNamedItem("Descricao").value
							apelido = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>" & desc_apelido & "&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'><b>" & apelido & "&nbsp;</b></td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' ORGAO DE ORIGEM
						'************************************************************
						if xmlCampos.nodeName = "ORGAO_ORIGEM" then
							codigo_orgao = trim(xmlCampos.attributes.getNamedItem("Codigo").value)
							tipo_orgao = trim(xmlCampos.attributes.getNamedItem("Tipo").value)
							desc_orgao = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							busca_orgao = RemoveTagMarcador(trim(xmlCampos.attributes.getNamedItem("Valor").value))
							desc_princ_orgao = replace(replace(replace(busca_orgao," ","_"),"&#39;","_#39;"),"'","\'")
							valor_orgao = trim(xmlCampos.attributes.getNamedItem("Valor").value)
							if (somente_detalhe = false) then
								select case tipo_orgao
									case "110"
										nome_funcao = "LinkBuscaOrgao"
										
										'TJ-RJ
										if (global_numero_serie = 4794) or (global_numero_serie = 5613) then
											hint = "Buscar todas legislações desta instituição"
										else
											hint = getTermo(global_idioma, 1559, "Buscar todos os registros desta instituição", 0)
										end if
										hint = hint & "..."
								End Select
								if (xmlCampos.attributes.getNamedItem("Link_Autoridades").value = "1") and (config_habilita_autoridades = 1) then 
									autInfo = "<td class='autImg'><div class='autImg' onClick=LinkAutInfo('detalhes',"&codigo_orgao&",'"&desc_princ_orgao&"',"&tipo_orgao&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1579, "Mostrar informações sobre este orgão de origem", 0)&"...'>&nbsp;</div></td>"
								else
									autInfo = ""
								end if                                
								temp = "<a class='link_classic2' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_orgao&",'"&desc_princ_orgao&"',"&iIndexSrv&");"">"&TrocaTagMarcador(replace(valor_orgao,chr(10),"<br/>")) &"</a>"
								autor = "<table class='autLink' style='border-spacing: 1px; padding: 0;'><tr><td id='autLinkTab' class='td_left_middle'>"&temp&"&nbsp;</td>"&autInfo&"</tr></table>"
							else
								autor = valor_orgao
							end if
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>"&desc_orgao&"</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & autor & "</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' ESFERA
						'************************************************************
						if xmlCampos.nodeName = "ESFERA" then
							desc_esfera = xmlCampos.attributes.getNamedItem("Descricao").value
							esfera = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>" & desc_esfera & "&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & esfera & "&nbsp;</td>"
							Response.Write "</tr>"
						end if					
						'*************************************
						'SITUAÇÃO DA LEGISLAÇÃO
						'*************************************
						if xmlCampos.nodeName = "SITUACAO_LEGISLACAO" then
							leg_situacao = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							leg_situacao_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							
							css_situacao = class_detalhe
							if (CStr(xmlCampos.attributes.getNamedItem("Destaque").value) = "1") then
								css_situacao = css_situacao&" td_leg_valor_detalhe"
							end if
							
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>" & leg_situacao_desc & "&nbsp;</td>"
							Response.Write "<td class='"&css_situacao&"'>" & leg_situacao & "&nbsp;</td>"
							Response.Write "</tr>"																									
						end if							
						'************************************************************
						' PAGINAS
						'************************************************************
						'if xmlCampos.nodeName = "PAGINAS" then
						'	desc_pag = xmlCampos.attributes.getNamedItem("Descricao").value
						'	pag = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
						'	Response.Write "<tr>"
						'	Response.Write "<td colspan='2' class='td_leg_descricao'>" & desc_pag & "&nbsp;</td>"
						'	Response.Write "<td class='"&class_detalhe&"'>" & pag & "&nbsp;</td>"
						'	Response.Write "</tr>"
						'end if
						'************************************************************
						' DATA DE ASSINATURA
						'************************************************************
						if xmlCampos.nodeName = "DATA_ASSINATURA" then
							desc_data_ass = xmlCampos.attributes.getNamedItem("Descricao").value
							data_ass = xmlCampos.attributes.getNamedItem("Valor").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>" & desc_data_ass & "&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & data_ass & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' DATA DE PUBLICAÇÃO
						'************************************************************
						if xmlCampos.nodeName = "DATA_PUBLICACAO" then
							desc_data_pub = xmlCampos.attributes.getNamedItem("Descricao").value
							data_pub = xmlCampos.attributes.getNamedItem("Valor").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>" & desc_data_pub & "&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & data_pub & "&nbsp;</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' RESPONSABILIDADE INTELECTUAL
						'************************************************************
						if xmlCampos.nodeName = "RESPONSABILIDADE_INTELECTUAL" then
							seq = 1
							autor = ""
							desc_responsabilidade = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>"&desc_responsabilidade&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								tipo_responsabilidade = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
								codigo_responsabilidade = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
								valor_responsabilidade = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								busca_responsabilidade = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								desc_princ_responsabilidade = replace(replace(replace(busca_responsabilidade," ","_"),"&#39;","_#39;"),"'","\'")
								seq_responsabilidade = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)
								if (somente_detalhe = false) then
									nome_funcao = "LinkBuscaAutor"
									
									'TJ-RJ
										if (global_numero_serie = 4794) or (global_numero_serie = 5613) then
											hint = "Buscar todas legislações deste autor"
										else
											hint = getTermo(global_idioma, 1558, "Buscar todos os registros deste autor", 0)
										end if
										hint = hint & "..."
									if (xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1") and (config_habilita_autoridades = 1) then
										'Adequação ABL - Exibir itens um em frente do outro
										if (global_numero_serie = 2372) or (global_numero_serie = 2635) then
											autInfo = "<div class='autImg div_imagem' id='autImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_responsabilidade&",'"&desc_princ_responsabilidade&"',"&tipo_responsabilidade&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1571, "Mostrar informações sobre este autor", 0)&"...'>&nbsp;</div>"
										else
											autInfo = "<td class='autImg'><div class='autImg' id='autImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_responsabilidade&",'"&desc_princ_responsabilidade&"',"&tipo_responsabilidade&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1571, "Mostrar informações sobre este autor", 0)&"...'>&nbsp;</div></td>"
										end if
									else
										autInfo = ""
									end if
									temp = seq_responsabilidade&" <a class='link_classic2' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_responsabilidade&",'"&desc_princ_responsabilidade&"',"&iIndexSrv&");"">"&replace(valor_responsabilidade,chr(10),"<br/>") &"</a>"
								else
									autInfo = ""
									temp = seq_responsabilidade & " " & valor_responsabilidade
								end if
                                identificadoresInfo = ""
                                For Each xmlIdentificadores in xmlSubCampos.childNodes                                    
	                                %><!--#include file ="identificadoresDetalhe.asp"--><%                                                                                                      
                                next 
                                autInfo = autInfo & identificadoresInfo
								if autor <> "" then						
									'Adequação ABL - Exibir itens um em frente do outro
									if (global_numero_serie = 2372) or (global_numero_serie = 2635) then
										autor = autor & "&nbsp;" & temp & "&nbsp;" & autInfo
									else
										autor = autor & "<table class='autLink' style='border-spacing: 1px; padding: 0;'><tr><td id='autLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & autInfo & "</tr></table>"
									end if
								else
									'Adequação ABL - Exibir itens um em frente do outro
									if (global_numero_serie = 2372) or (global_numero_serie = 2635) then
										autor = temp & "&nbsp;" & autInfo
									else
										autor = "<table class='autLink' style='border-spacing: 1px; padding: 0;'><tr><td id='autLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & autInfo & "</tr></table>"
									end if
								end if
								seq = seq + 1
							Next
							Response.Write "<td class='"&class_detalhe&"'>" & autor & "</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' PROJETOS DE LEI
						'************************************************************
						if xmlCampos.nodeName = "PROJETOS_LEI" then
							desc_projetos_lei = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>"&desc_projetos_lei&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
							Response.Write "</tr>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' PROJETO DE LEI
								'************************************************************
								desc_projeto_lei = xmlSubCampos.attributes.getNamedItem("Descricao").value
								projeto_lei = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								Response.Write "<tr>"
								Response.Write "<td class='centro' style=' height: 16px; width: 19px'><div class='joinbottom icon_20'></div></td>"
								Response.Write "<td class='td_leg_descricao'>" & replace(desc_projeto_lei,chr(10),"<br/>") & "&nbsp;</td>"
								Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & replace(projeto_lei,chr(10),"<br/>") & "&nbsp;</div></td>"
								Response.Write "</tr>"
							Next
						end if
                        '************************************************************
						' PROCESSOS
						'************************************************************
						if xmlCampos.nodeName = "PROCESSOS" then
                            processo = ""
                            desc_processo = xmlCampos.attributes.getNamedItem("Descricao").value
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' PROCESSO
								'************************************************************
                                if (processo <> "") then
                                    processo = processo & ", "
                                end if
								processo = processo & TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
							Next
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>" & desc_processo & "&nbsp;</td>"
							Response.Write "<td class='" & class_detalhe & "'>" & processo & "</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' ASSUNTOS
						'************************************************************
						if xmlCampos.nodeName = "ASSUNTOS" then
							seq = 1
							assunto = ""
							desc_assunto = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
								codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
								valor_assunto = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								busca_assunto = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								desc_princ_assunto = replace(replace(replace(busca_assunto," ","_"),"&#39;","_#39;"),"'","\'")
								seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)
								if (somente_detalhe = false) then
									nome_funcao = "LinkBuscaAssunto"
									'TJ-RJ
									if (global_numero_serie = 4794) or (global_numero_serie = 5613) then
										hint = "Buscar todas legislações deste assunto"
									else
										hint = getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)
									end if
									
									if (trim(hint) <> "") then
										hint = hint & "..."
									end if
									
									if (xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1") and (config_habilita_autoridades = 1) then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assInfo = "<div class='assuntoImg div_imagem' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div>"
										else
											assInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
										end if
									else
										assInfo = ""
									end if
									temp = seq_assunto&" <a class='link_custom' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_assunto&",'"&desc_princ_assunto&"',"&iIndexSrv&");"">"&replace(valor_assunto,chr(10),"<br/>") &"</a>"
								else
									assInfo = ""
									temp = seq_assunto & " " & valor_assunto
								end if
                                identificadoresInfo = ""
                                For Each xmlIdentificadores in xmlSubCampos.childNodes                                    
	                                %><!--#include file ="identificadoresDetalhe.asp"--><%                                                                                                      
                                next 
                                assInfo = assInfo & identificadoresInfo
								if assunto <> "" then
									'Adequação ABL - Exibir itens um em frente do outro
									if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
										assunto = assunto & "&nbsp;" & temp & "&nbsp;" & assInfo
									else
										assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0;''><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
									end if
								else
									'Adequação ABL - Exibir itens um em frente do outro
									if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
										assunto = assunto & temp & "&nbsp;" & assInfo
									else
										assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0;'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
									end if
								end if
								seq = seq + 1
							Next
							Response.Write "<td class='"&class_detalhe&"'>" & assunto & "</td>"
							Response.Write "</tr>"
						end if
						'************************************************************
						' EMENTA
						'************************************************************
						if xmlCampos.nodeName = "ABSTRACT" then
							nome_funcao = "LinkBuscaEmenta"
							hint = getTermo(global_idioma, 1605, "Exibir texto completo", 0)&"..."
							modo = "em"
							Mais = trim(xmlCampos.attributes.getNamedItem("Mais").value)
							Ementa = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							Desc_Ementa = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao' >"&Desc_Ementa&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & Ementa
							if cStr(Mais) = "1" then
								Response.Write "&nbsp;...</div><div class='direita'><span class='span_imagem div_imagem_right_3 icon_9 mais-png '></span><a class='link_classic2' title='"&hint&"' onClick="&nome_funcao&"("&codigo_obra&",'"&modo&"',"&iIndexSrv&")  href=""javascript:"&nome_funcao&"("&codigo_obra&",'"&modo&"');"">"&getTermo(global_idioma, 1621, "Ler mais", 0)&"</a>&nbsp;&nbsp;</div></td>"
                            else
								Response.Write "</div></td>"
							end if
							Response.Write "</tr>"					
						end if
						'************************************************************
						' OBSERVAÇÕES
						'************************************************************
						if xmlCampos.nodeName = "OBSERVACOES" then
							nome_funcao = "LinkBuscaLegObservacoes"
							hint = getTermo(global_idioma, 1605, "Exibir texto completo", 0)&"..."
							modo = "obsleg"
							Mais = trim(xmlCampos.attributes.getNamedItem("Mais").value)
							Observacoes = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							Desc_Observacoes = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>"&Desc_Observacoes&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & Observacoes
							if cStr(Mais) = "1" then
								Response.Write "&nbsp;...</div><div class='direita'><span class='span_imagem div_imagem_right_3 icon_9 mais-png '></span><a class='link_classic2' title='"&hint&"' onClick="&nome_funcao&"("&codigo_obra&",'"&modo&"',"&iIndexSrv&")  href=""javascript:"&nome_funcao&"("&codigo_obra&",'"&modo&"');"">"&getTermo(global_idioma, 1621, "Ler mais", 0)&"</a>&nbsp;&nbsp;</div></td>"
							else
								Response.Write "</div></td>"
							end if
							Response.Write "</tr>"					
						end if
						'************************************************************
						' TEXTO INTEGRAL
						'************************************************************
						if xmlCampos.nodeName = "TEXTO_INTEGRAL" then
							nome_funcao = "LinkBuscaTextoInt"
							hint = getTermo(global_idioma, 1605, "Exibir texto completo", 0)&"..."
							modo = "ti"
							Mais = trim(xmlCampos.attributes.getNamedItem("Mais").value)
							Texto = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							Desc_Texto = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>"&Desc_Texto&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & Texto
							if cStr(Mais) = "1" then
								Tamanho = trim(xmlCampos.attributes.getNamedItem("Tamanho").value)
								FuncTextoIntegral = nome_funcao&"("&codigo_obra&",'"&modo&"',"&iIndexSrv&","&Tamanho&")"
								Response.Write "&nbsp;...</div><div class='direita'><span class='span_imagem div_imagem_right_3 icon_9 mais-png '></span><a class='link_classic2' title='"&hint&"' onClick="""&FuncTextoIntegral&""" href=""javascript:"&FuncTextoIntegral&";"">"&getTermo(global_idioma, 1621, "Ler mais", 0)&"</a>&nbsp;&nbsp;</div></td>"
								if (Request.QueryString("integra") = "1") then
									ScriptIntegra = FuncTextoIntegral
								end if
							else
								Response.Write "</div></td>"
							end if
							Response.Write "</tr>"					
						end if
						'************************************************************
						' PUBLICAÇÃO
						'************************************************************
						if xmlCampos.nodeName = "PUBLICACAO" then
							desc_publicacao = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>"&desc_publicacao&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
							Response.Write "</tr>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' PUBLICAÇÕES
								'************************************************************
								desc_publicacao = xmlSubCampos.attributes.getNamedItem("Descricao").value
								publicacao = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								Response.Write "<tr>"
								Response.Write "<td class='centro' style='height: 16px; width: 19px'><div class='joinbottom icon_20'></div></td>"
								Response.Write "<td class='td_leg_descricao'>"&desc_publicacao&"&nbsp;</td>"
								Response.Write "<td class='"&class_detalhe&"'><div class='justificado'>" & replace(publicacao,chr(10),"<br/>") & "&nbsp;</div></td>"
								Response.Write "</tr>"
							Next
						end if
						'************************************************************
						' ALTERAÇÕES
						'************************************************************
						if xmlCampos.nodeName = "ALTERACOES" then
							alteracoes = ""
							desc_alteracao = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>"&desc_alteracao&"&nbsp;</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' ALTERAÇÕES
								'************************************************************
								cod_alteracao = xmlSubCampos.attributes.getNamedItem("Codigo").value
								desc_alteracao = xmlSubCampos.attributes.getNamedItem("Descricao").value
								alteracao = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								if alteracoes <> "" then
									alteracoes = alteracoes & "<br/><a class='link_classic2' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1,"&Trim(cod_alteracao)&",1,'leg',2);"">"&replace(alteracao,chr(10),"<br/>")&"</a>"
								else
									alteracoes = "<a class='link_classic2' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1,"&Trim(cod_alteracao)&",1,'leg',2);"">"&replace(alteracao,chr(10),"<br/>")&"</a>"
								end if
							Next
							Response.Write "<td class='"&class_detalhe&"'>" & alteracoes
							Response.Write "</td></tr>"
						end if
						'************************************************************
						' CORRELAÇÕES
						'************************************************************
						if xmlCampos.nodeName = "CORRELACOES" then
							correlacoes = ""
							desc_correlacao = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao'>"&desc_correlacao&"&nbsp;</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' CORREÇÕES
								'************************************************************
								cod_correlacao = xmlSubCampos.attributes.getNamedItem("Codigo").value
								correlacao = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
                                if cod_correlacao = 0 then
									if correlacoes <> "" then
										correlacoes = correlacoes & "<br/>"&replace(correlacao,chr(10),"<br/>")
									else
										correlacoes = replace(correlacao,chr(10),"<br/>")
									end if
                                else
									if correlacoes <> "" then
										correlacoes = correlacoes & "<br/><a class='link_classic2' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1,"&Trim(cod_correlacao)&",1,'leg',2);"">"&replace(correlacao,chr(10),"<br/>")&"</a>"
									else
										correlacoes = "<a class='link_classic2' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1,"&Trim(cod_correlacao)&",1,'leg',2);"">"&replace(correlacao,chr(10),"<br/>")&"</a>"
									end if
                                end if
							Next
							Response.Write "<td class='"&class_detalhe&"'>" & correlacoes
							Response.Write "</td></tr>"
						end if
						'************************************************************
						' SITES RELACIONADOS
						'************************************************************
						if xmlCampos.nodeName = "SITES" then
							sites = ""
							desc_sites = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao' >"&desc_sites&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>&nbsp;</td>"
							Response.Write "</tr>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'************************************************************
								' SITES
								'************************************************************
								desc_site = xmlSubCampos.attributes.getNamedItem("Descricao").value
								site_c = ""
								For Each xmlItem In xmlSubCampos.childNodes
									codMidia = xmlItem.attributes.getNamedItem("Codigo").value
									plataforma = xmlItem.attributes.getNamedItem("Plataforma").value
									site = TrocaTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
									site_url = RemoveTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
                                    dspace = xmlItem.attributes.getNamedItem("DSpace").value
									if site_c <> "" then
										site_c = site_c & "<br/>"
									end if
									if (plataforma <> "") then
										site_c = site_c & "<a class='link_classic2' href='javascript:abreMidiaEspecifica(" & Tipo & "," & codigo_obra & "," & codMidia & "," & iIndexSrv & ")'>" & site & "</a>"
									else
										site_c = site_c & "<a class='link_classic2' href='" & site_url & "' target='_blank' onclick='ContarAcesso("&codigo_obra&","&codMidia&",1,"&dspace&");'>" & site & "</a>"
									end if
								Next	
								
								Response.Write "<tr>"
								Response.Write "<td class='td_detalhe_descricao_imagem'>"
								Response.Write "<div class='joinbottom icon_20'></div>"
								Response.Write "</td>"
								Response.Write "<td class='td_detalhe_descricao_recuo'>"
								Response.Write desc_site&"&nbsp;"
								Response.Write "</td>"
								Response.Write "<td class='"&class_detalhe&"'>" & site_c & "&nbsp;</td>"
								Response.Write "</tr>"
							Next
						end if
						'************************************************************
						' URL
						'************************************************************
						if xmlCampos.nodeName = "URL" then
                            url = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_url = xmlCampos.attributes.getNamedItem("Descricao").value
							Response.Write "<tr>"
							Response.Write "<td colspan='2' class='td_leg_descricao' >"&desc_url&"&nbsp;</td>"
							Response.Write "<td class='"&class_detalhe&"'>" & url & "index.asp?codigo_sophia=" & codigo_obra & "</td>"
							Response.Write "</tr>"
							
							'Adequação TJ-RJ - Exibe link para o texto integral
							if (global_numero_serie = 4794) then
								Response.Write "<tr>"
								Response.Write "<td colspan='2' class='td_leg_descricao'>Link para texto integral&nbsp;</td>"
								Response.Write "<td class='"&class_detalhe&"'>" & url & "index.asp?codigo_sophia=" & codigo_obra & "&integra=1</td>"
								Response.Write "</tr>"
							end if
						end if
					Next
                    'Para resolver o problema no IE, quando utilizado colspan é ignorado o width da coluna
					Response.Write "<tr><td style='width: 19px;'></td><td style='width: 141px'></td><td></td></tr>"
					Response.Write "</table>"
					if global_exibe_capa = 1 then
						Response.Write "</td></tr>"
						Response.Write "</table>"
					end if
				'************************************************************
				' SERVIÇOS
				'************************************************************
				elseif xmlPNome.nodeName = "SERVICOS" then
					For Each xmlServicos In xmlPNome.childNodes
						'*************************************
						'SERVIÇOS HABILITADOS PARA A LEGISLAÇÃO
						'*************************************
						if xmlServicos.nodeName = "LINK_SELECIONAR" AND (somente_detalhe = false) and (global_esconde_menu = 0) then
							serv_selecao = true
						end if
						if xmlServicos.nodeName = "LINK_MIDIAS" then
							serv_midias = true
							serv_midias_desc = xmlServicos.attributes.getNamedItem("Descricao").value
							serv_midias_mus = xmlServicos.attributes.getNamedItem("Audio").value
						end if
					Next
				end if
			Next
		end if
	end if	

	'*************************************
	' ANALITICAS DE OBRAS (Somente BNBD)
	'*************************************
	if (global_numero_serie = 5516) AND (serv_analitica = true) then
		if (global_exibe_capa = 1) then
			styleTable = "margin-left: 114px; width: 826px;"
			style = " padding-left: 114px;"
		else
			styleTable = "width: 100%;"
			style = ""
		end if

		Response.Write "</td></tr>"
		Response.Write "<tr><td class='centro' style='padding-bottom: 6px;" & style & "'><br/><b>" & getTermo(global_idioma, 53, "Analíticas", 0) & "</b></td></tr>"
		Response.Write "<tr><td>"
		Response.Write "<table style='padding: 0; border-spacing: 1px; background-color: #999999;" & styleTable & "'>"
		Response.Write "<tr style='height: 20px'>"
		Response.Write "<td class='esquerda td_tabelas_titulo'>&nbsp;" & getTermo(global_idioma, 177, "Título", 0) & "</td>"
		Response.Write "<td class='centro td_tabelas_titulo' style='width: 100px'>" & getTermo(global_idioma, 4218, "Paginação", 0) & "</td>"
		Response.Write "</tr>"

		Set ROService = ROServer.CreateService("Web_Consulta")
		sXmlAnaliticas = ROService.TitulosAnaliticas(codigo_obra)
		Set ROService = nothing

		sequencial = 0

		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml sXmlAnaliticas
		Set xmlRoot = xmlDoc.documentElement
	
		For Each xmlCampos In xmlRoot.childNodes

			sequencial = sequencial + 1
						
			if ((sequencial mod 2) > 0) then
				'---------- IMPAR
				fontcolor = "black"
				td_class = "td_tabelas_valor1"
				link_class = "link_serv"
			else
				'------------ PAR
				fontcolor= "#000000"
				td_class = "td_tabelas_valor2"
				link_class = "link_serv"
			end if

			'************************************************************
			' Analitica
			'************************************************************
			if xmlCampos.nodeName = "TITULO" then
				codigo_analitica = xmlCampos.attributes.getNamedItem("CODIGO").value
				titulo_analitica = xmlCampos.attributes.getNamedItem("TITULO").value
				paginacao_analitica = xmlCampos.attributes.getNamedItem("PAGINACAO").value

				sTitAnalitica = "<a class='" & link_class & "' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,0,"&codigo_analitica&","&pagina&",'analitica',3);"">"&titulo_analitica&"</a>"

				Response.Write "<tr style='height: 19px'>"
				Response.Write "<td class='esquerda " & td_class & "'>&nbsp;" & sTitAnalitica & "</td>"
				Response.Write "<td class='centro " & td_class & "'>" & paginacao_analitica & "</td>"
				Response.Write "</tr>"
			end if
		Next
		
		Response.Write "</table><br />"
	end if
	
	'*************************************
	'AVALIAÇÃO (RANKING)
	'*************************************
	htmlAvaliacao = ""

	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	
	xml_AvaliacaoOnline = ROService.DetalhesAvaliacaoOnLine(codigo_obra)
	
	TrataErros(1)
	
	SET ROService = nothing
	
	if left(xml_AvaliacaoOnline,5) = "<?xml" then
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml xml_AvaliacaoOnline
		Set xmlRoot = xmlDoc.documentElement
		'************************************************************
		' INFORMAÇÃO DO TITULO
		'************************************************************
		if xmlRoot.nodeName = "avaliacoes" then

			mediaAvaliacao = xmlRoot.attributes.getNamedItem("media").value
			quantidadeAvaliacao = xmlRoot.attributes.getNamedItem("quantidade").value
		end if
	end if
	
	'if xmlCampos.nodeName = "avaliacao" then
							
		'Verifica se avaliação está habilitada
		if (global_avaliacao_online = 1) then
							
			htmlAvaliacao = "<div style='float: left; height: 25px; padding-top: 3px;'>" & ImprimeEstrelas(mediaAvaliacao,false) & "</div>"	

			if quantidadeAvaliacao = 0 then

				if global_avaliacao_autenticada = 1 then
					if Session("Logado")= "sim" then
						linkAvaliacao = "<a class='link_classic2' title='" & getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "' style='cursor:pointer; color: #00395B;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"&CodigoUsuario=" & Session("codigo_usuario") & "','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
						htmlAvaliacao = htmlAvaliacao & "<div style='float: left; margin-left: 10px; height: 25px; padding-top:3px;'>" & linkAvaliacao & getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "</a></div>"
					else
						linkAvaliacao = "<a class='link_classic2' title='" & getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "' style='cursor:pointer;color: #00395B;' onclick=""abreLogin('avaliacao_votar','"&getTermo(global_idioma, 6689, "Avaliação on-line", 0)&"','&codigo_obra="&codigo&"',true,true);"">"
						htmlAvaliacao = htmlAvaliacao  & "<div style='float: left; margin-left: 10px; height: 25px; padding-top:3px;'>" &  linkAvaliacao & getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "</a></div>"
					end if
				else
					linkAvaliacao = "<a class='link_classic2' title='" & getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"&CodigoUsuario=0','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
					htmlAvaliacao = htmlAvaliacao  & "<div style='float: left; margin-left: 10px; height: 25px; padding-top:3px;'>" &   linkAvaliacao & getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "</a></div>"
				end if
			elseif quantidadeAvaliacao = 1 then
				if global_avaliacao_autenticada = 1 then
					if Session("Logado")= "sim" then
						linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"&CodigoUsuario=" & Session("codigo_usuario") & "','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
						linkMediaAvaliacao = "<a class='link_classic2' title='" & getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) &"' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
						htmlAvaliacao = htmlAvaliacao & "<div style='float: left; margin-left: 10px; height: 25px; padding-top:3px;'>" & linkMediaAvaliacao & getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</a></div>"
					else 
						linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;color: #00395B;' onclick=""abreLogin('avaliacao_votar','"&getTermo(global_idioma, 6689, "Avaliação on-line", 0)&"','&codigo_obra="&codigo&"',true,true);"">"
						linkMediaAvaliacao = "<a class='link_classic2' title='" & getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) &"' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
						htmlAvaliacao = htmlAvaliacao &  "<div style='float: left; margin-left: 10px; height: 25px; padding-top:3px;'>" & linkMediaAvaliacao & getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</a></div>"
					end if
				else
					linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"&CodigoUsuario=0','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
					linkMediaAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) &"' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
					htmlAvaliacao = htmlAvaliacao &  "<div style='float: left; margin-left: 10px; height: 25px; padding-top:3px;'>" & linkMediaAvaliacao & getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</a></div>"
				end if
			else
				if global_avaliacao_autenticada = 1 then
					if Session("Logado")= "sim" then
						linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"&CodigoUsuario=" & Session("codigo_usuario") & "','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
						linkMediaAvaliacao = "<a class='link_classic2' title='"& Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) &"' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
						htmlAvaliacao = htmlAvaliacao & "<div style='float: left; margin-left: 10px; height: 25px; padding-top:3px;'>" & linkMediaAvaliacao & Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</a></div>"
					else
						linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;color: #00395B;' onclick=""abreLogin('avaliacao_votar','"&getTermo(global_idioma, 6689, "Avaliação on-line", 0)&"','&codigo_obra="&codigo&"',true,true);"">"
						linkMediaAvaliacao = "<a class='link_classic2' title='"& Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) &"' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
						htmlAvaliacao = htmlAvaliacao & "<div style='float: left; margin-left: 10px; height: 25px; padding-top:3px;'>" & linkMediaAvaliacao & Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</a></div>"
					end if
				else
					linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"&CodigoUsuario=0','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
					linkMediaAvaliacao = "<a class='link_classic2' title='"& Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) &"' style='cursor:pointer;color: #00395B;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
					htmlAvaliacao = htmlAvaliacao & "<div style='float: left; margin-left: 10px; height: 25px; padding-top:3px;'>" & linkMediaAvaliacao & Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</a></div>"
				end if
			end if
		end if

	'end if
	'***************************************************************
	Set xmlDoc = nothing
	Set xmlRoot = nothing
else
	Response.Write getTermo(global_idioma, 1273, "Nenhuma informação encontrada", 0)
end if

'*************************************
'REDES SOCIAIS
'*************************************
						
divRedesSociais = ""

linkRedesSociais = Session("baseUrl") & "info.asp?c="&codigo_obra
if global_facebook_curtir = 1 then
	
	' Utilização do Curtir do Facebook via API
	if (global_facebook_iframe <> 1) then
	
		divRedesSociais = divRedesSociais & "<div class='p-facebook' style='float: left; margin-left: 10px; height: 25px; width: 100px; overflow:hidden;visibility:hidden;'><fb:like href="&linkRedesSociais&" send='false' show_faces='false' layout='button_count' ></fb:like></div>"

	' Utilização do Curtir do Facebook via iFRAME
	else

		divRedesSociais = divRedesSociais & "<div style='display: inline-table; margin-left: 10px; height: 25px;'>" & _
			"<iframe src=""//www.facebook.com/plugins/like.php?href=" & Server.URLEncode(linkRedesSociais) & "&amp;" & _
            "layout=button_count&amp;width=100&amp;show_faces=false&amp;font&amp;colorscheme=light&amp;action=like&amp;height=20"" scrolling=""no"" " & _
            "frameborder=""0"" style=""border:none; overflow:hidden; width:100px; height:20px;"" allowTransparency=""true""></iframe></div>"

	end if

end if

if global_twitter = 1 then
								
	divRedesSociais = divRedesSociais & "<div style='float: left; margin-left: 10px; height: 25px;' ><a href='https://twitter.com/share' class='twitter-share-button' data-counturl='"&linkRedesSociais&"' data-url='"&linkRedesSociais&"' data-text='"&getTermoHtml(global_idioma, 7211, "Confira", 0)&": '>Tweet</a><script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0],p=/^http:/.test(d.location)?'http':'https';if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src=p+'://platform.twitter.com/widgets.js';fjs.parentNode.insertBefore(js,fjs);}}(document, 'script', 'twitter-wjs');</script></div>"

end if

'************************************************************
divRankingRedesSociais = ""

if htmlAvaliacao <> "" and divRedesSociais = "" then
	divRankingRedesSociais = "<div style='height: 15px; width:270px; margin: auto; padding-left:40px; margin-top: 15px'>" & htmlAvaliacao & "</div>"
elseif htmlAvaliacao = "" and divRedesSociais <> "" then
	divRankingRedesSociais = "<div style='height: 15px; width:270px; margin: auto; padding-left:40px; margin-top: 15px'>" & divRedesSociais & "</div>"
elseif htmlAvaliacao <> "" and divRedesSociais <> "" then
	divRankingRedesSociais = "<div style='height: 15px; width:495px; margin: auto; padding-left:40px; margin-top: 15px'>" & htmlAvaliacao & divRedesSociais & "</div>"
end if

Response.Write "</td></tr>"
Response.Write "</table>"
Response.Write divRankingRedesSociais & "<br /><br />"

'**************************************************************************
' MONTA OS SERVIÇOS DISPONIVEIS PARA O TITULO
'**************************************************************************
if serv_selecao = true OR serv_reserva = true OR serv_aquisicao = true OR serv_midias = true OR _
   serv_refbib = true OR serv_bibcomp = true OR serv_analitica = true then
	Response.Write "<table style='display: inline-block;'><tr>"
	'*************************************
	'SELECIONAR
	'*************************************
	if (veio_de <> "sels") and (serv_selecao = true) then
		s_sel = "<a class='link_serv' title='"&getTermo(global_idioma, 1350, "Enviar para minha seleção", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/selecionar.asp?veio_de=grid&codigo="&codigo_obra&"&tipo="&Tipo&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 963, "Minha seleção", 0)&"',320,210,false,false);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-selecao-b'></span>"&getTermo(global_idioma, 128, "Selecionar", 0)&"</a>"
		Response.Write "<td class='centro td_detalhe_serv'>"&s_sel&"</td>"
	end if
	'*************************************
	'Favoritos
	'*************************************
	if (veio_de <> "favoritos") then 
        if global_hab_servicos = 1 then
		    if Session("Logado") = "sim" then
                logado = "true"
			    if config_multi_servbib = 1 then
				    if CStr(Session("Servidor_Logado")) = CStr(iIndexSrv) then
					    sLink_Favoritos = "<a class='link_serv' title='"&getTermo(global_idioma, 8318, "Clique aqui para salvar os títulos selecionados nos favoritos.", 0)&"' href=""javascript:salvar_registro_favoritos("&codigo_obra&","&iIndexSrv&");"">"
                        sEstadoIconeFavorito = "transparent-icon icon-small-star-b"
				    else
                        sEstadoIconeFavorito = "icon-small-star-x"
					    sLink_Favoritos = "<a class='link_serv_custom' onclick=""abreLogin('favoritos','"&getTermo(global_idioma, 8317, "Salvar favoritos", 0)&"','&codigosSelecionados="&codigo_obra&"&servidor=parent.hiddenFrame.iSrvCorrente_MySel',false,true);"">"
				    end if
			    else
				    sEstadoIconeFavorito = "transparent-icon icon-small-star-b"
                    sLink_Favoritos = "<a class='link_serv' title='"&getTermo(global_idioma, 8318, "Clique aqui para salvar os títulos selecionados nos favoritos.", 0)&"' href=""javascript:salvar_registro_favoritos("&codigo_obra&","&iIndexSrv&");"">"
			    end if
		    else
                sEstadoIconeFavorito = "icon-small-star-x"
			    sLink_Favoritos = "<a class='link_serv_custom' onclick=""abreLogin('favoritos','"&getTermo(global_idioma, 8317, "Salvar favoritos", 0)&"','&codigosSelecionados="&codigo_obra&"&servidor=parent.hiddenFrame.iSrvCorrente_MySel',false,true);"">"
		    end if	

			sLink_Favoritos = sLink_Favoritos & "<span class='span_imagem div_imagem_right_3 icon_16 "&sEstadoIconeFavorito&"'></span>"&getTermo(global_idioma, 8317, "Salvar favoritos", 0)&"</a>"
			Response.Write "<td class='centro td_detalhe_serv'>"&sLink_Favoritos&"</td>"
		end if
	end if
	'*************************************
	'MÍDIAS
	'*************************************
	if serv_midias = true then
		if global_numero_serie = 2635 then 'ABL_BRG
			s_midia = "<a class='link_serv' title='Visualizar doc. digital' style='cursor:pointer;' onclick=""abrePopup('asp/midia.asp?tipo="&Tipo&"&codigo="&codigo_obra&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','Doc. digital',500,490,true,true);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-digital'></span>&nbsp;Doc. digital&nbsp;</a>"
		else
			if serv_midias_mus = "1" then
				s_midia = "<a class='link_serv' title='"&getTermo(global_idioma, 1622, "Ouvir áudio", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/midia.asp?tipo="&Tipo&"&codigo="&codigo_obra&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&serv_midias_desc&"',500,490,true,true);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-audio'></span>&nbsp;"&serv_midias_desc&"&nbsp;</a>"
			else
				s_midia = "<a class='link_serv' title='"&getTermo(global_idioma, 1622, "Visualizar conteúdo digital", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/midia.asp?tipo="&Tipo&"&codigo="&codigo_obra&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 7517, "Conteúdo digital", 0)&"',500,490,true,true);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-digital-b'></span>&nbsp;"&serv_midias_desc&"&nbsp;</a>"
			end if
		end if
		
		Response.Write "<td class='centro td_detalhe_serv'>"&s_midia&"</td>"
	end if
	'*************************************
	'ANALITICAS
	'*************************************
	if (serv_analitica = true) AND (global_numero_serie <> 5516) then
		s_ana = "<a class='link_serv' style='cursor:pointer;' title='"&getTermo(global_idioma, 1402, "Exibir analíticas", 0)&"' onclick=""abrePopup('asp/artigos.asp?obra="&codigo_obra&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"&Servidor="&iIndexSrv&"','"&serv_analitica_desc&"', 965, 450, false, true);"");><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-document-b'></span>&nbsp;"&serv_analitica_desc&"&nbsp;</a>"
		Response.Write "<td class='centro td_detalhe_serv'>"&s_ana&"</td>"
	end if
	'*************************************
	'REFERENCIA BIBLIOGRAFICA
	'*************************************
	if serv_refbib = true then
		s_refbib = "<a class='link_serv' style='cursor:pointer;' title='"&getTermo(global_idioma, 1403, "Exibir referência bibliográfica", 0)&"' onclick=exibeRefbib('"&Tipo&"',"&codigo_obra&","&iIndexSrv&");><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-book-b'></span>&nbsp;"&getTermo(global_idioma, 5507, "Referência", 0)&"&nbsp;</a>"
		Response.Write "<td class='centro td_detalhe_serv'>"&s_refbib&"</td>"
	end if
	'*************************************
	'BIBLIOGRAFIA COMPLEMENTAR
	'*************************************
	if serv_bibcomp = true then
		s_bibcomp = "<a href='#bibComp' class='link_serv' style='cursor:pointer;' title='"&getTermo(global_idioma, 1405, "Exibir veja também", 0)&"' onclick=exibeBibcomp("&codigo_obra&");><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-baloonadd'></span>&nbsp;"&getTermo(global_idioma, 6379, "Veja também", 0)&"&nbsp;</a>"
		Response.Write "<td class='centro td_detalhe_serv' style='width: 130px'>"&s_bibcomp&"</td>"
	end if
	'*************************************
	'RESERVA
	'*************************************
	if serv_reserva = true and (Session("usuario_externo") = false) then
		if Session("Logado") = "sim" then
			if config_multi_servbib = 1 then
				if CStr(Session("Servidor_Logado")) = CStr(iIndexSrv) then
					s_res = "<a class='link_serv' title='"&getTermo(global_idioma, 1459, "Reservar", 0)&"' style='cursor:pointer;' onclick=""if(parent.hiddenFrame.modo_busca='telacheia'){parent.hiddenFrame.modo_busca='rapida';}abrePopup('asp/reserva.asp?veio_de=busca_principal&codigo_obra="&codigo_obra&"&tipo_obra="&Tipo&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 348, "Reserva", 0)&"',380,400,false,false);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-flag-b'></span>"&getTermo(global_idioma, 1459, "Reservar", 0)&"</a>"
				else
					s_res = "<a class='link_serv' title='"&getTermo(global_idioma, 1407, "Efetue o login para reservar", 0)&"' style='cursor:pointer;' onclick=""abreLogin('reserva','"&getTermo(global_idioma, 348, "Reserva", 0)&"','&codigo_obra="&codigo_obra&"&tipo_obra="&Tipo&"&servidor="&iIndexSrv&"',false,true);""><span class='span_imagem div_imagem_right_3 icon_16 icon-small-flag-x removerBordas'></span>"&getTermo(global_idioma, 1459, "Reservar", 0)&"</a>"
				end if
			else
				s_res = "<a class='link_serv' title='"&getTermo(global_idioma, 1459, "Reservar", 0)&"' style='cursor:pointer;' onclick=""if(parent.hiddenFrame.modo_busca='telacheia'){parent.hiddenFrame.modo_busca='rapida';}abrePopup('asp/reserva.asp?veio_de=busca_principal&codigo_obra="&codigo_obra&"&tipo_obra="&Tipo&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 348, "Reserva", 0)&"',380,400,false,false);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-flag-b'></span>"&getTermo(global_idioma, 1459, "Reservar", 0)&"</a>"
			end if
		else
			s_res = "<a class='link_serv' title='"&getTermo(global_idioma, 1407, "Efetue o login para reservar", 0)&"' style='cursor:pointer;' onclick=""abreLogin('reserva','"&getTermo(global_idioma, 348, "Reserva", 0)&"','&codigo_obra="&codigo_obra&"&tipo_obra="&Tipo&"&servidor="&iIndexSrv&"',false,true);""><span class='span_imagem div_imagem_right_3 icon_16 icon-small-flag-x removerBordas'></span>"&getTermo(global_idioma, 1459, "Reservar", 0)&"</a>"
		end if
		Response.Write "<td class='centro td_detalhe_serv'>"&s_res&"</td>"
	end if
	'*************************************
	'AQUISIÇÃO
	'*************************************
	if serv_aquisicao = true then
		if Session("Logado") = "sim" and Session("sgt_aquisicoes") = 1 and (Session("usuario_externo") = false) then
			if config_multi_servbib = 1 then
				if CStr(Session("Servidor_Logado")) = CStr(iIndexSrv) then
					s_aq = "<a class='link_serv' title='"&getTermo(global_idioma, 1408, "Sugerir aquisição", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/nova_sugestao.asp?codigo_obra="&codigo_obra&"&modo_busca=parent.hiddenFrame.modo_busca&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 1321, "Sugestões", 0)&"',710,425,false,false);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-lamp-b'></span>"&getTermo(global_idioma, 184, "Aquisição", 0)&"</a>"
					Response.Write "<td class='centro td_detalhe_serv'>"&s_aq&"</td>"
				end if
			else
				s_aq = "<a class='link_serv' title='"&getTermo(global_idioma, 1408, "Sugerir aquisição", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/nova_sugestao.asp?codigo_obra="&codigo_obra&"&modo_busca=parent.hiddenFrame.modo_busca&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 1321, "Sugestões", 0)&"',710,425,false,false);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-lamp-b'></span>"&getTermo(global_idioma, 184, "Aquisição", 0)&"</a>"
				Response.Write "<td class='centro td_detalhe_serv'>"&s_aq&"</td>"
			end if
		end if
	end if
	Response.Write "</tr></table>"
end if

%>
<br/><br />
<div id="div_refbib"></div>
<tr><td class="centro" colspan="5" style="padding-top: 7px;">
<%
if ExibeEx = true then
	if (CStr(Tipo) = "2") then
		'**************************************************************************
		' MONTA O GRID DE FONTES DA LEGISLAÇÃO
		'**************************************************************************
		%><!-- #include file="grid_leg_per.asp" --><%
	elseif (CStr(Tipo) <> "3") and (CStr(Tipo) <> "20") and (CStr(Tipo) <> "21") then
		'**************************************************************************
		' MONTA O GRID DE EXEMPLARES
		'**************************************************************************
		%><!-- #include file="grid_exemplares.asp" --><%
	end if
end if 
%>
</td>
</tr>
</table>
</td>
</tr>
</table>

<!--#include file ="identificadoresDetalheImagem.asp"-->

<%
	if (Request.QueryString("refresh_popup") = "") then
		if (Request.QueryString("midiaext") <> "") then
			codMidia = Request.QueryString("midiaext")
			redirecttw = Request.QueryString("twri")
			Response.Write "<script>"
			Response.Write "function abreMidia() { "
			Response.Write "abrePopup('asp/midia.asp?tipo="&Tipo&"&codigo="&codigo_obra&"&iIndexSrv="&iIndexSrv&"&codMidia="&codMidia&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"&twri="&redirecttw&"','"&getTermo(global_idioma, 7517, "Conteúdo digital", 0)&"',500,490,true,true);"
			Response.Write "} "
			Response.Write "</script>"		
		else
			if (ScriptIntegra <> "") then
				Response.Write "<script language='javascript' type='text/javascript'>"
				Response.Write "function abreTextoIntegral() { "
				Response.Write ScriptIntegra & "; "
				Response.Write "} "
				Response.Write "</script>"
			end if
		end if
	end if
%>

