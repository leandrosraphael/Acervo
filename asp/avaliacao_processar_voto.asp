<!DOCTYPE>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
    
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge" />
	<title></title>
	<!-- #include file="../config.asp" -->
	<!-- #include file="../idiomas/idiomas.asp" -->
	<!-- #include file="../libasp/header.asp" -->
	<!-- #include file="../libasp/funcoes.asp" -->
	
    <link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" /> 
    <link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />
	<% if config_css_estatico = 1 then %>
		<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
        
	<% else %>
		<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
	<% end if %>
	
	<link href="../inc/avaliacao_processar_voto.css" rel="stylesheet" type="text/css" /> 
    <link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
    <link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
    
    <%
	
		iCodigoObra  = request.QueryString("CodigoObra")
		iNotaObra = Request("nota_avaliacao")
		iCodigoUsuario = request.QueryString("CodigoUsuario")

		Set ROService = ROServer.CreateService("Web_Consulta")	
		Set objComputarVoto = ROServer.CreateComplexType("TMensagem")

		Set objComputarVoto = ROService.GravarAvaliacaoOnLine(iCodigoObra,iNotaObra,iCodigoUsuario,global_idioma)

		bRetorno = objComputarVoto.Result
		sMsg 	 = objComputarVoto.sMsg

		Set ROService = nothing
		TrataErros(1)

		mensagem = sMsg
	 %>

	<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
	<script type="text/javascript">
		function SetRefreshURL() {
			if (parent.parent.hiddenFrame != null) {
				Hiddenfrm = parent.parent.hiddenFrame;
				Hiddenfrm.popup_refresh = true;
			}
		}
	</script>

</head>
<body class="popup" onload="parent.fechaLoadingPopup();">

	<div id="mensagem">
	
		<%=mensagem%>

	</div>
	
	<div class="div-BotaoFechar">

		<%
			'Limpar o cache das últimas aquisições sempre que for realizada uma avaliação
			DefinirCache "resultadoUltAquisicoes_" & global_numero_serie & "_" & iIndexSrv & "_" & repositorio_institucional, ""
			DefinirCache "sXMLPaginasUltAquisicoes_" & global_numero_serie & "_" & iIndexSrv & "_" & repositorio_institucional, ""
			DefinirCache "sXMLFichasUltAquisicoes_" & global_numero_serie & "_" & iIndexSrv & "_" & repositorio_institucional, ""
		%>

		<input class="button_termo" type="button" onclick="javascript:SetRefreshURL();parent.fechaPopup('asp');" value="<%=getTermo(global_idioma, 220, "Fechar", 0)%>" id="fechar"/>
	</div>

</body>
</html>
