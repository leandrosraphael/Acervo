<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
<%
sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<title>::<%=getTermo(global_idioma, 167, "Excluir", 0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />

<% if config_css_estatico = 1 then %>
	<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
<% else %>
	<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
<% end if %>

<link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />

<script src="../scripts/jquery/jquery-3.2.1.min.js"></script>

    <script type="text/javascript" src="../scripts/jquery/jquery-ui.min.js"></script>
	<script type="text/javascript" src="../scripts/jquery.multiselect.js"></script>
	<script type="text/javascript" src="../scripts/combo.init.js"></script>
	<link href="../scripts/css/estilo_idioma.css" rel="stylesheet" type="text/css" />
	<link href="../scripts/css/jquery-ui-1.8.18.datepicker.css" rel="stylesheet" />

	<link href="../scripts/css/jquery.multiselect.css" rel="stylesheet" />
	<link href="../scripts/css/multiselect-custom.css" rel="stylesheet" />
	<link href="../scripts/css/pubtype-icons.css" rel="stylesheet" type="text/css" />


	<script type="text/javascript" src="../scripts/favoritos.js"></script>

</head>
<body class="popup">
<script type="text/javascript">parent.fechaLoadingPopup();</script>

<br />
<div class="centro">
	<div style="height: 90px">
    
<%
	idioma = IntQueryString("idioma", 0)

	iIndexSrv = IntQueryString("servidor", 1)
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%

	codigoFavorito = IntQueryString("codigoFavorito", 0)
	
    if IsNumeric(codigoFavorito) then
		Set ROService = ROServer.CreateService("Web_Consulta")
		Set Msg = ROService.ExcluirConteudoFavorito(codigoFavorito, idioma)
		Set ROService = nothing
		TrataErros(1)
		%></div><br /><%
		if Msg.sMsg <> "" then
            sMens = replace(Msg.sMsg,"[QUEBRA]","<br />")
			Response.Write sMens
			Response.Write "<input type='button' onclick=""parent.fechaPopup('asp');parent.atualiza_Lista_Favoritos("&iIndexSrv&", 0);"" value='"&getTermo(idioma, 220, "Fechar", 0)&"' />"
		else
			Response.Write "<script type='text/javascript'>parent.fechaPopup('asp');parent.atualiza_Lista_Favoritos("&iIndexSrv&", 0);</script>"
		end if
	end if 
	%>
</div>
</body>
</html>