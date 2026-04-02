<%
	sDiretorioArq = "asp"
    nao_imprime_variaveis_globais = "1"
%>

<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<%
    iIndexSrv = IntQueryString("iIndexSrv", 1)
	grupo = request.QueryString("grupo")
%>

<!-- #include file="../libasp/updChannelProperty.asp" -->
<%
	Set ROService = ROServer.CreateService("Web_Consulta")

    if (grupo = "material") then
	    Resultado = ROService.BuscarDadosAceleradorMaterial
        TermoGrupo = getTermo(global_idioma, 175, "Material", 0)
    else
        Resultado = ROService.BuscarDadosAceleradorRepositorio
        TermoGrupo = getTermo(global_idioma, 270, "Repositório digital", 0)
    end if

	Set ROService = nothing

	Totalizadores = ""

	if (left(Resultado, 5) = "<?xml") then
	    Set xmlDoc = CreateObject("Microsoft.xmldom")
	    xmlDoc.async = False
	    xmlDoc.loadxml Resultado

		Set xmlRoot = xmlDoc.documentElement
		
		if (xmlRoot.nodeName = "TOTALIZADORES") then
            Totalizadores = Totalizadores & "<div class='titulo-acelerador-busca linha'><span>" & TermoGrupo & "</span></div>"
            Totalizadores = Totalizadores & "<div class='corpo-acelerador-busca'>"        

            For Each xmlGrupo In xmlRoot.childNodes
                Totalizadores = Totalizadores & _
                "<div class='linha' onclick='RetornarResultadoDeBuscaDoAcelerador(" & xmlGrupo.attributes.getNamedItem("CODIGO").value & ",""" & grupo & """)'> " & _
                  "<div>" & _
                    "<span>" & xmlGrupo.attributes.getNamedItem("GRUPO").value & "</span>" & _
                  "</div>" & _
                  "<div>" & _
                    "<span>" & xmlGrupo.attributes.getNamedItem("QUANTIDADE").value & "</span>" & _
                  "</div>" & _
                "</div>"
            next

            Totalizadores = Totalizadores & "</div>"
		end if

		Set xmlDoc = nothing
	end if

    Response.write Totalizadores
%>