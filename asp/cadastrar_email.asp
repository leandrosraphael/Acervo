<% Response.Buffer=True %>
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
</head>
<body class="popup">
<%
	if IsNumeric(Request.QueryString("usuario")) AND IsNumeric(Request.QueryString("codigo_obra")) then
		email = Request.Form("email")
		if len(replace(email," ","")) = 0 OR Instr(email,"@") = 0 OR len(trim(email)) = 0 then
			Response.Write "<p class='centro'>"&getTermo(global_idioma, 1317, "O e-mail não é válido.", 0)
			Response.Write "<br /><a href='javascript:history.go(-1);'>"&getTermo(global_idioma, 1386, "voltar", 2)&"</a></p>"
		else
			On Error Resume Next
			SET ROService = ROServer.CreateService("Web_Consulta")
			ROService.AtualizaEmail trim(Request.QueryString("usuario")), trim(email)
			SET ROService = nothing
			TrataErros(1)
%>
			<script type='text/javascript'>alert('<%=getTermo(global_idioma, 1387, "E-mail cadastrado com sucesso.", 0)%>\n<%=getTermo(global_idioma, 1388, "Clique em OK para continuar com a reserva.", 0)%>');</script>
			<form name="login" method="post" action="reserva.asp">
			<input type="hidden" name="codigo_obra" value="<%=trim(Request.QueryString("codigo_obra"))%>" />
			<input type="hidden" name="tipo_obra" value="<%=trim(Request.QueryString("tipo_obra"))%>" />
			<input type="hidden" name="ano" value="<%=trim(Request.QueryString("ano"))%>" />
			<input type="hidden" name="volume" value="<%=trim(Request.QueryString("volume"))%>" />
			<input type="hidden" name="edicao" value="<%=trim(Request.QueryString("edicao"))%>" />
			<input type="hidden" name="suporte" value="<%=trim(Request.QueryString("suporte"))%>" />
			<input type="hidden" name="biblioteca" value="<%=trim(Request.QueryString("biblioteca"))%>" />
			<input type="hidden" name="servidor" value="<%=IntQueryString("Servidor",1)%>" />
			<input type="hidden" name="veio_de_old" value="<%=trim(Request.QueryString("veio_de"))%>" />
			<input type="hidden" name="veio_de" value="email" />
			<script type='text/javascript'>document.login.submit();</script>
			</form>
			<%		    
			Response.Write "</p>"
		end if
	else
		Response.Write "<p class='centro'>"&getTermo(global_idioma, 100, "Código de usuário inválido.", 0)&"</p>"
	end if
%>
</body>
</html>