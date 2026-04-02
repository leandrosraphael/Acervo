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
<title>::<%=getTermo(global_idioma, 129, "E-mail", 0)%></title>
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
<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
<script type="text/javascript">
<!--

function validateEmailUsu(email,frm)
{
	var tam_email = email.length;
	if (tam_email < 4) {
		var invalido = 1;
	} else {
		email 			= email.replace(",",";");
		var array_email	= email.split(";");
		var email_num	= 0;
		var invalido 	= 0;
		while (email_num < array_email.length)
		{
			array_email[email_num] = Trim(array_email[email_num]);
			if (typeof(array_email[email_num]) != "string") {
				invalido+=1;
			} else if (!array_email[email_num].match(/^[A-Za-z0-9]+([_.-][A-Za-z0-9]+)*@[A-Za-z0-9]+([_.-][A-Za-z0-9]+)*\.[A-Za-z0-9]{2,4}$/)) {
				invalido+=1;
			}
			email_num+=1;
		}
	}

	if (invalido > 0) 
	{
		alert('<%=getTermo(global_idioma, 1317, "O e-mail não é válido.", 0)%>');
		return false;
	}
	else
	{
		frm.submit();
	}
}

-->
</script>
</head>
<% 
	iIndexSrv = Session("Servidor_Logado")
	
	if iIndexSrv = "" then
		iIndexSrv = 1
	end if

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%

	if len(trim(Session("codigo_usuario"))) = 0 then
		Response.Write "<script type='text/javascript'>alert('"&getTermo(global_idioma, 1474, "Sessão expirada", 0)&".\n"&getTermo(global_idioma, 4097, "Favor logar novamente.", 0)&"');parent.fechaPopup();</script>"
		Response.End()
	end if
	
	if request.querystring("acao") = "alterar" then	
		codUsu = Request.QueryString("usuario")
		sEmail = Request.Form("email")
	
		On Error Resume Next
		SET ROService = ROServer.CreateService("Web_Consulta")
		ROService.AtualizaEmail codUsu, sEmail
		SET ROService = nothing
		TrataErros(1)

		Response.Write "<script type='text/javascript'>parent.document.location='../index.asp?content=dsi&modo_busca=" & GetModo_Busca & "&iBanner="&global_tipo_banner&"&iIdioma=" & global_idioma & "';</script>"
		Response.Write "<script type='text/javascript'>parent.fechaPopup();</script>"	
	else
%>
<body class="popup">
<script type="text/javascript">parent.fechaLoadingPopup();</script>
<% 
	'------------------------
	'Busca e-mail atual do usuário
	Set ROService = ROServer.CreateService("Web_Consulta")	
	sEmail = ROService.EmailUsuario(Session("codigo_usuario"))
	Set ROService = nothing
	TrataErros(1)

	Response.Write "<p class='centro'><br/><br/><br/>"
	Response.Write getTermo(global_idioma, 3710, "Por favor, informe seu e-mail.", 0)
	Response.Write "<br/><br/><br/>"
	Response.Write "<form name='frm_email' method='post' action='alterar_email.asp?acao=alterar&modo_busca=" & GetModo_Busca & "&usuario=" & Session("codigo_usuario") & "&iBanner="&global_tipo_banner&"&iIdioma=" & global_idioma & "'>"
	Response.Write "<input class='input_email' type='text' name='email' value=" & sEmail & ">"
	Response.Write "<br/><br/><br/><br/>"
	Response.Write "<input type='button' value='" & getTermo(global_idioma, 510, "Alterar", 0) & "' onclick='return validateEmailUsu(document.frm_email.email.value,document.frm_email)'>"
	Response.Write "&nbsp;&nbsp;<input type='button' value='" & getTermo(global_idioma, 5, "Cancelar", 0) & "' onclick='parent.fechaPopup()'>"
	Response.Write "</form>"
	Response.Write "</p>"
%>
</body>
<% end if %>
</html>