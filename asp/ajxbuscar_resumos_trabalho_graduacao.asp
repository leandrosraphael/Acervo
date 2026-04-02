<%
	sDiretorioArq = "asp"
    nao_imprime_variaveis_globais = "1"
%>

<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<%
    iIndexSrv = IntQueryString("iIndexSrv", 1)
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%

	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	Set FiltroTeseDissertacao = ROServer.CreateComplexType("TFiltroTeseDissertacao")
	Set Resultado = ROServer.CreateComplexTupe("TString")
	
	anoInicial = request.QueryString("anoInicial")

	if (anoInicial <> "") then
		anoInicial = ClNg(anoInicial)
	else
		anoInicial = 0
	end if
	FiltroTeseDissertacao.AnoInicial = anoInicial

	anoFinal = request.QueryString("anoFinal")
	if (anoFinal <> "") then
		anoFinal = ClNg(anoFinal)
	else
		anoFinal = 0
	end if
	FiltroTeseDissertacao.AnoFinal = anoFinal

	tipoFiltro = CInt(request.QueryString("tipoFiltro"))

	FiltroTeseDissertacao.TipoFiltroAno = tipoFiltro

	Set Resultado = ROService.buscarResumoTrabalhoGraduacao(FiltroTeseDissertacao)

	Set ROService = nothing
	Set FiltroTeseDissertacao = nothing

	Totalizadores = ""
	Links = ""

    if (left(Resultado.Param0, 5) = "<?xml") then
	    Set xmlDoc = CreateObject("Microsoft.xmldom")
	    xmlDoc.async = False
	    xmlDoc.loadxml Resultado.Param0
		
		Set xmlRoot = xmlDoc.documentElement

		iTotalCursos = 0
		
		for Each xmlConsulta in xmlRoot.childNodes
			if (xmlConsulta.nodeName = "GRADUACAO") then
				termoTitulo = getTermo(global_idioma, 9350, "Trabalhos de Graduação", 0)
				Facetas =	"<div class='area-faceta-consulta background-faceta'>"  & _
							"	<div class='tab-busca-facetada'>" & _
							"		<p>"&termoTitulo&"</p>" & _
							"	</div>"

				Facetas = Facetas & _
							"<div id='conteudoTrabalhoGraduacao'>" & _
							"	<ul class='ul-faceta lista-faceta-consulta'>" & _
							"		<li class='titulo-faceta expandida'>" & _
							"			<span>Graduação</span>" & _
							"			<span class='transparent-icon span_imagem icon_16 icon-small-down-b'></span>" & _
							"		</li>"

				For Each xmlFaceta In xmlConsulta.childNodes
					if (xmlFaceta.nodeName = "FACETA") then
						descricaoFaceta = getTermo(global_idioma, xmlFaceta.attributes.getNamedItem("TERMO").value, xmlFaceta.attributes.getNamedItem("DESCRICAO").value, 0) & _
							" (" & xmlFaceta.attributes.getNamedItem("TOTAL").value & ")"
						Facetas = Facetas & _
							"<li class='item-faceta-consulta'>" & descricaoFaceta & "</li>"
					end if
				next

				Facetas = Facetas & "</ul></div></div>"
			end if
		next

		Set xmlDoc = nothing
	end if

	if (left(Resultado.Param1, 5) = "<?xml") then
	    Set xmlDoc = CreateObject("Microsoft.xmldom")
	    xmlDoc.async = False
	    xmlDoc.loadxml Resultado.Param1
		
		Set xmlRoot = xmlDoc.documentElement
		for Each xmlConsulta in xmlRoot.childNodes
			if (xmlConsulta.nodeName = "LINKS") then
				iTotalCursos = xmlConsulta.attributes.getNamedItem("TOTAL_CURSOS_GRADUACAO").value
				
				Links =	"<div class='area-links-pesquisa'>" & _
						"	<div class='titulo-detalhe-pesquisa'>" & _
						"		<p><span>" & iTotalCursos & "</span>Trabalhos de Graduação</p>" & _
						"</div>" & _
						"<div>" & _
						"	<ul>"

				For Each xmlLinks In xmlConsulta.childNodes
					if (xmlLinks.nodeName = "LINKS") then
						codigoCurso = xmlLinks.attributes.getNamedItem("CODIGO_CURSO").value
						linkBuscarTitulos = "pesquisarTitulosTrabalhoGraduacao("&codigoCurso&","&anoInicial&","&anoFinal&","&tipoFiltro&");"
						Links = Links & _
							"<li class='li-negrito'><a class='link_custom_negrito' href='#' onclick='"&linkBuscarTitulos&"'>" & xmlLinks.attributes.getNamedItem("DESCRICAO").value & "</a></li>"
					end if
				next

				Links = Links & "</ul></div></div>"
			end if
		next

        Set xmlRoot = nothing
	end if
	
	if (iTotalCursos = 0) then
		html_resultado_busca = "<div class='centro'><span style='color: red'>"&getTermo(global_idioma, 1341, "Nenhum registro encontrado.", 0)&"</div>"
	else
		html_resultado_busca = Facetas & Links & _
			"<div class='nota-rodape-pesquisa'>" & _
				"	<p style='display: inline-block;'>O curso ""Engenharia Civil-Aeronáutica"" até 2006 era denominado ""Engenharia de Infra-Estrutura Aeronáutica"".</p>" & _
			"</div>"
	end if

	Response.write html_resultado_busca
	%>
	<script type="text/javascript">
		global_frame.codigoCurso = 0;
		global_frame.codigoPrograma = 0;
		global_frame.codigoMaterial = 0;
		global_frame.codigoArea = 0;
		global_frame.tipoBuscaRegistro = 1;
		global_frame.anoInicial = <%= anoInicial %>;
		global_frame.anoFinal = <%= anoFinal %>;
		global_frame.tipoFiltro = <%= tipoFiltro %>;
	</script>