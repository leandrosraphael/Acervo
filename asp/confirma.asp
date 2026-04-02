<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<title>::Login</title>
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
<script type='text/javascript'>
<!--
function validaTecla(campo, event) { 
	var key; 
	var tecla; 
	CheckTAB=true;
	tecla= event.keyCode; 
	key = String.fromCharCode( tecla); 
	//alert('aqui');
	if ( tecla == 13 )
	{
		document.login.submit();		
	}
}
-->
</script>
</head>
<body class="popup" style="margin-left: 5px; margin-top: 5px;" onload="document.login.codigo.focus();">
<p class="centro">
<%if Request.Form("sub_confirma") = "sim" then
	Senha_Usuario = Request.Form("senha")
	if Senha_Usuario = Session("Senha_Usuario") then
		Session("confirma") = "sim"
	else
		Session("confirma") = "nao"
	end if
	Response.Write("<script type='text/javascript'>window.close();</script>")	
	
else%>

	<form name="confirma" method="post" action="troca_senha.asp" onSubmit="">
	<table style="width: 225px;">
	<tr>
	<td colspan="2" class="centro">
    <%
		msg_senha = getTermo(1455, "Para maior segurança você deve confirmar sua senha %satual%s para realizar esta operação", 0)
		msg_senha = Format(msg_senha, "<b>|</b>")&":"
		Response.Write msg_senha
	%>
	<br /><br /></td></tr>
	<tr>
	<td><%=getTermo(global_idioma, 2, "Senha", 0)%></td>
	<td class="esquerda"><input type="password" name="senha" size="20" maxlength="6" onKeyPress="return validaTecla(this, event)"></td>
	</tr>
	<tr>
	<td colspan="2" class="centro"><br /><br /><input type="button" onClick="submit()" value="Confirma" id="button1" name="button1"></td>
	</tr>
	</table>
	<input type="hidden" name="sub_troca_senha" value="sim"/>
	</form>

<%end if%>
</p>
</body>
</html>