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
	<b>Recepción de mensajes de la biblioteca en <i>Facebook</i> - Instrucciones</b>
</p>
<div style="text-align: left;margin-left:10px;">
	<br />
	<br />
	Al registrar un nombre de usuario de <i>Facebook</i> en el Opac Web de SophiA, se habilita la recepción de e-mails de la biblioteca en la bandeja de entrada de la cuenta de <i>Facebook</i> indicada.
	<br />
	Para que eso sea posible, por favor, siga las siguientes instrucciones:
	<br />
	<br />
	<b>1. Registrar el nombre de usuario de <i>Facebook</i> en el área <i>Informaciones personales</i> del Opac Web.</b>
	<br />
	<br />
	<div style="margin-left:10px;">
		1.1. Accede al link <a href="https://www.facebook.com/settings?tab=account&section=username" target="_blank">https://www.facebook.com/settings?tab=account&amp;section=username</a>;
		<br />
		<br />
		1.2. Entra a la cuenta de <i>Facebook</i>;
		<br />
		<br />
		1.3. Copiar el nombre del campo  <b>Nombre de usuario</b>:
		<br />
		<br />
		<div class='help-facebook-username-es-es'></div>
		<br />
		<br />
		1.4. Pegar este nombre en el campo <b><i>Facebook</i></b> de las <b>Informaciones personales</b> en el Opac Web; 
		<br />
		<br />
		1.5. Confirmar las modificaciones.
		<br />
		<br />
		<div class='help-facebook-info-es-es'></div>
	</div>
	<br />
	<br />
	<b>2. Añadir el usuario de la biblioteca como amigo en <i>Facebook</i>.</b>
	<br />
	<br />
	<div style="margin-left:10px;">
		<% if global_facebook_perfil <> "" then %>
			2.1. Accede al link <a href="<%= global_facebook_perfil %>" target="_blank"><%= global_facebook_perfil %></a>;
		<% else %>
			2.1. Accede al perfil de <i>Facebook</i> de la biblioteca;
		<% end if %>
		<br />
		<br />
		2.2. Entra a la cuenta de <i>Facebook</i>;
		<br />
		<br />
		2.3. Haz clic en el botón  "+ Añadir a mis amigos":
		<br />
		<br />
		<div class='help-facebook-add-friend-es-es'></div>
		<br />
		<br />
		2.4. Confirmación del sistema.
		<br />
		<br />
		Observaciones: hasta que el usuario de la biblioteca no sea añadido como amigo en <i>Facebook</i>, los mensajes de la biblioteca sólo podrán ser recibidos en la bandeja de entrada "Otros", que puede ser accedida según se indica en la ventana siguiente:
		<br />
		<br />
		<div class='help-facebook-other-inbox-es-es'></div>
	</div>
</div>
</body>
</html>