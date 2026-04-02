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
<title></title>
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

<script type='text/javascript' src='../scripts/ajxSeries.js'></script>
<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>

</head>
<body class="popup" style="padding-top: 6px;">
<script type="text/javascript">parent.fechaLoadingPopup();</script>
<p class="centro">
	<b>Recebimento de mensagens via <i>Facebook</i> - Instruções</b>
</p>
<div style="text-align: left;margin-left:10px;">
	<br />
	<br />
	Ao cadastrar o seu nome de usuário do <i>Facebook</i>, você habilitará o recebimento de e-mails da biblioteca na caixa de entrada da sua conta do <i>Facebook</i>.
	<br />
	Para que isso seja possível, por favor, siga os seguintes procedimentos:
	<br />
	<br />
	<b>1. Cadastre o seu nome de usuário do <i>Facebook</i> em <i>Informações pessoais</i> do Terminal Web.</b>
	<br />
	<br />
	<div style="margin-left:10px;">
		1.1. Acesse o link <a href="https://www.facebook.com/settings?tab=account&section=username" target="_blank">https://www.facebook.com/settings?tab=account&amp;section=username</a>;
		<br />
		<br />
		1.2. Efetue o <i>login</i> com sua conta do <i>Facebook</i> (caso ainda não tenha feito);
		<br />
		<br />
		1.3. Copie o nome descrito no campo <b>Nome de usuário</b>:
		<br />
		<br />
		<div class='help-facebook-username'></div>
		<br />
		<br />
		1.4. Efetue o <i>login</i> no Terminal Web;
		<br />
		<br />
		1.5. Acesse Serviços - Informações pessoais e clique em Alterar;
		<br />
		<br />
		1.6. Cole as informações no campo <b><i>Facebook</i></b>;
		<br />
		<br />
		1.7. Confirme as alterações.
		<br />
		<br />
		<div class='help-facebook-info'></div>
	</div>
	<br />
	<br />
	<b>2. Adicione o usuário da biblioteca como seu amigo no <i>Facebook</i>.</b>
	<br />
	<br />
	<div style="margin-left:10px;">
		<% if global_facebook_perfil <> "" then %>
			2.1. Acesse o link <a href="<%= global_facebook_perfil %>" target="_blank"><%= global_facebook_perfil %></a>;
		<% else %>
			2.1. Acesse o perfil de <i>Facebook</i> da biblioteca;
		<% end if %>
		<br />
		<br />
		2.2. Efetue o <i>login</i> com sua conta do <i>Facebook</i> (caso ainda não tenha feito);
		<br />
		<br />
		2.3. Clique no botão "+ Adicionar aos amigos":
		<br />
		<br />
		<div class='help-facebook-add-friend'></div>
		<br />
		<br />
		2.4. Aguarde a confirmação de amizade pela Biblioteca.
		<br />
		<br />
		Obs.: Enquanto o usuário da biblioteca não é adicionado como amigo no <i>Facebook</i>, as mensagens da biblioteca poderão ser recebidas na caixa de entrada "<b>Outras</b>", que pode ser acessada conforme a imagem a seguir:
		<br />
		<br />
		<div class='help-facebook-other-inbox'></div>
	</div>
</div>
</body>
</html>