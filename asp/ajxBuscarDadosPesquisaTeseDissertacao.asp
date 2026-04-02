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
	
%>

<!-- #include file="../libasp/updChannelProperty.asp" -->

<%
	Set ROService = ROServer.CreateService("Web_Consulta")
	Set FiltroTeseDissertacao = ROServer.CreateComplexType("TFiltroTeseDissertacao")
	Set Resultado = ROServer.CreateComplexTupe("TString")
	
	Idioma = ClNg(request.QueryString("idioma"))
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
	material = CInt(request.QueryString("material"))

	FiltroTeseDissertacao.TipoFiltroAno = tipoFiltro
	FiltroTeseDissertacao.Material = material
	FiltroTeseDissertacao.Idioma = Idioma

	Set Resultado = ROService.BuscarDadosPesquisaTeseDissertacao(FiltroTeseDissertacao)

	Set ROService = nothing
	Set FiltroTeseDissertacao = nothing

	Totalizadores = ""
	Links = ""

	'XML dos totalizadores
	if (left(Resultado.Param0, 5) = "<?xml") then
	    Set xmlDoc = CreateObject("Microsoft.xmldom")
	    xmlDoc.async = False
	    xmlDoc.loadxml Resultado.Param0

		Set xmlRoot = xmlDoc.documentElement
		
		if (xmlRoot.nodeName = "TOTALIZADORES") then
			Totalizadores =	_
						"<div class='area-faceta-consulta background-faceta'>"  & _
						"	<div class='tab-busca-facetada'>" & _
						"		<p>" & getTermo(Idioma, 2194, "Teses", 0) & "/" & getTermo(Idioma, 9396, "Dissertações", 0) & "</p>" & _
						"	</div>"

			tituloFacetaMargemTop = ""

			For Each xmlMaterial In xmlRoot.childNodes
				Totalizadores = Totalizadores & _
							"<div class='conteudoTeseDissertacao'>" & _
							"	<ul class='ul-faceta lista-faceta-consulta'>" & _
							"		<li class='titulo-faceta expandida" & tituloFacetaMargemTop & "'>" & _
							"			<span>" & getTermo(Idioma, xmlMaterial.attributes.getNamedItem("ID_TERMO").value, xmlMaterial.attributes.getNamedItem("DESCRICAO").value, 0) & " (" & _ 
												  xmlMaterial.attributes.getNamedItem("TOTAL").value & ")" & _
							"			</span>" & _
							"			<span class='transparent-icon span_imagem icon_16 icon-small-down-b'></span>" & _
							"		</li>"

				tituloFacetaMargemTop = " titulo-faceta-margin-top"

				For Each xmlTotalizador In xmlMaterial.childNodes
					
					Totalizadores = Totalizadores & _
						"<li class='item-faceta-consulta'>" & _
							getTermo(Idioma, xmlTotalizador.attributes.getNamedItem("ID_TERMO").value, xmlTotalizador.attributes.getNamedItem("DESCRICAO").value, 0) & _
							" (" & xmlTotalizador.attributes.getNamedItem("TOTAL").value & ")" & _
						"</li>"
				next

				Totalizadores = Totalizadores & "</ul></div>"
			next
		
			Totalizadores = Totalizadores & "</div>"
		end if

		Set xmlDoc = nothing
	end if

	EncontrouRegistro = false

	'XML dos dados de pesquisa
	if (left(Resultado.Param1, 5) = "<?xml") then
	    Set xmlDoc = CreateObject("Microsoft.xmldom")
	    xmlDoc.async = False
	    xmlDoc.loadxml Resultado.Param1

		Set xmlRoot = xmlDoc.documentElement

		if (xmlRoot.nodeName = "PESQUISA") then
			Links =	"<div class='area-links-pesquisa-doisniveis'>" & _
					"	<div class='titulo-detalhe-pesquisa'>" & _
					"		<p class='nota-pesquisa'>" & _
					"			<span>" & xmlRoot.attributes.getNamedItem("TOTAL").value & "</span>" & _
								getTermo(idioma, 2194, "Teses", 0) & "/" & getTermo(idioma, 9396, "Dissertações", 0) & _
					"		</p>"
			if (global_numero_serie = 3156) then
				Links =	Links & "<p class='nota-pesquisa nota-ita'>(*) " & getTermo(idioma, 0, "Programa/Área descontinuado", 0) & "</p>"
			end if
			Links = Links &	_
				"	</div>" & _
				"	<div>" & _
					"		<ul>"
			
			For Each xmlPrograma in xmlRoot.childNodes
				codigoPrograma = xmlPrograma.attributes.getNamedItem("CODIGO").value
				if (codigoPrograma = 0) then
					codigoPrograma = "-1"
				end if
				linkBuscarRegistrosArea = "pesquisarTitulosTesePorPrograma("&material&","&codigoPrograma&","&anoInicial&","&anoFinal&","&tipoFiltro&");"
				Links = Links & _
					"	<li class='li-negrito'><a class='link_custom_negrito' href='#' onclick='"&linkBuscarRegistrosArea&"'>" & getTermo(idioma, xmlPrograma.attributes.getNamedItem("ID_TERMO").value, xmlPrograma.attributes.getNamedItem("DESCRICAO").value, 0) & " (" & _
											xmlPrograma.attributes.getNamedItem("TOTAL").value & ")</a>" & _
					"		<ul class='lista-com-borda'>"
				
				For Each xmlArea In xmlPrograma.childNodes
					codigoArea = xmlArea.attributes.getNamedItem("CODIGO").value
					if (codigoArea = 0) then
						codigoArea = "-1"
					end if
				    linkBuscarRegistrosAreaPrograma = "pesquisarTitulosTesePorProgramaArea("&material&","&codigoPrograma&","&codigoArea&","&anoInicial&","&anoFinal&","&tipoFiltro&");"
					Links = Links & _
						"	<li>" & _
						"		<a class='link_custom' href='#' onclick='"&linkBuscarRegistrosAreaPrograma&"'>" & getTermo(idioma, xmlArea.attributes.getNamedItem("ID_TERMO").value, xmlArea.attributes.getNamedItem("DESCRICAO").value, 0) & " (" & _
												xmlArea.attributes.getNamedItem("TOTAL").value & ")</a>" & _
						"	</li>"
				next
			
				Links = Links & "</ul></li>"

				EncontrouRegistro = true
			next

			Links = Links & "</ul></div>"
		end if

		Set xmlDoc = nothing
	end if
	
	Set Resultado = nothing

	if (EncontrouRegistro) then
		Response.write Totalizadores & Links
		if (global_numero_serie = 3156) then
			Response.write "<div class='nota-rodape-pesquisa'><p style='display: inline-block;'>(*) " & getTermo(idioma, 0, "Programa/Área descontinuado", 0) & "</p></div>"
		end if
	else
		Response.write "<div class='centro'><span style='color: red'>" & getTermo(idioma, 1341, "Nenhum registro encontrado.", 0) & "</div>"
	end if	
%>
<script type="text/javascript">
	global_frame.codigoCurso = 0;
	global_frame.codigoPrograma = 0;
	global_frame.tipoBuscaRegistro = 2;
	global_frame.codigoMaterial = <%= material %>;
	global_frame.anoInicial = <%= anoInicial %>;
	global_frame.anoFinal = <%= anoFinal %>;
	global_frame.tipoFiltro = <%= tipoFiltro %>;
</script>