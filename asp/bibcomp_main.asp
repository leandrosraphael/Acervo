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
</head>
<body class="popup" onload="parent.fechaLoadingPopup();">
	<%
	iIndexSrv = IntQueryString("servidor", 1)
	codigo = request.QueryString("codigo")
	%>
    <p class='centro'>
		<br />
		<div class="esquerda" style="margin-bottom: 4px; margin-left: 12px;"><b><%=getTermo(global_idioma, 205, "Assuntos", 0)%></b></div>
		<iframe class="if_bibcomp" frameborder="0" src="bibcomp.asp?cod=<%=codigo%>&tipo=Assuntos&servidor=<%=iIndexSrv%>&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>" scrolling="yes"></iframe>
		<br />
		<br />
		<div class="esquerda" style="margin-bottom: 4px; margin-left: 12px;">&nbsp;<b><%=getTermo(global_idioma, 1176, "Autores", 0)%></b></div>
		<iframe class="if_bibcomp" frameborder="0" src="bibcomp.asp?cod=<%=codigo%>&tipo=Autores&servidor=<%=iIndexSrv%>&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>" scrolling="yes"></iframe>
		<br />
		<br />
	</p>
</body>
</html>