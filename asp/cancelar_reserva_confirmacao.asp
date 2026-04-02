<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
<% sDiretorioArq="asp" %>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<title>::<%=getTermo(global_idioma, 963, "Minha seleção", 0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />

<% 
Response.Write "<link href='../inc/estilo_padrao.css' rel='stylesheet' type='text/css' />"
Response.Write "<link href='../inc/imagem_padrao.css' rel='stylesheet' type='text/css' />"
if config_css_estatico = 1 then 
	Response.Write "<link href='../inc/estilo.css' rel='stylesheet' type='text/css' />"
else 
	Response.Write "<link href='../libasp/estilo.asp?idioma="&global_idioma&"' rel='stylesheet' type='text/css' />"
end if 
Response.Write "<link href='../inc/contraste.css' rel='stylesheet' type='text/css' />"
Response.Write "<link href='../inc/imagem_contraste.css' rel='stylesheet' type='text/css' />"
%>
<script src="../scripts/jquery/jquery-3.2.1.min.js"></script>

<script type="text/javascript" src="../scripts/jquery/jquery-ui.min.js"></script>
<script type="text/javascript" src="../scripts/modal.js"></script>
<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
</head>
<%
global_idioma = request.QueryString("iIdioma")
mensagem = getTermo(global_idioma, 4707, "Confirma a exclusão da reserva selecionada?", 0)
codigo = request.QueryString("codigo")
digital = request.QueryString("digital")  
%>

<body class="popup">
	<script type="text/javascript">parent.fechaLoadingPopup();</script>
	<br />
	<div class="centro" style="width: 94%; margin: auto">
		<br /><%Response.Write mensagem%><br /><br /><br />
		<a class='link_serv' href="javascript:ExcluiReserva(<%=codigo%>,<%=digital%>);"><%=getTermo(global_idioma, 783, "Sim", 0)%></a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a class='link_serv' href="javascript:parent.fechaPopup('asp');"><%=getTermo(global_idioma, 4054, "Não", 0)%></a>
	</div>
</body>
</html>