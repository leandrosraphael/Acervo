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

<script type='text/javascript' src='../scripts/ajxSeries.js'></script>
<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>

</head>
<body class="popup" style="padding-top: 6px;">
<script type="text/javascript">parent.fechaLoadingPopup();</script>
<p class="centro">
	<b>Recepció de missatges de la biblioteca amb <i>Facebook</i> - Instruccions </b>
</p>
<div style="text-align: left;margin-left:10px;">
	<br />
	<br />
	En registrar un nom d'usuari de <i>Facebook</i> en l'Opac Web de SophiA, s'habilita la recepció d'e-mails de la biblioteca a la safata de entrada del compte de Facebook indicat.
	<br />
	Per què això sigui possible, seguiu sisplau les següents instruccions:
	<br />
	<br />
	<b>1. Registreu el nom d'usuari de <i>Facebook</i> en l'àrea <i>Informacions personals</i> de l'Opac Web.</b>
	<br />
	<br />
	<div style="margin-left:10px;">
		1.1. Accediu al link <a href="https://www.facebook.com/settings?tab=account&section=username" target="_blank">https://www.facebook.com/settings?tab=account&amp;section=username</a>;
		<br />
		<br />
		1.2. Entreu al compte de <i>Facebook</i>;
		<br />
		<br />
		1.3. Copieu el nom del camp <b>Nom d'usuari</b>: 
		<br />
		<br />
		<div class='help-facebook-username-es-ca'></div>
		<br />
		<br />
		1.4. Enganxeu aquest nom en el camp <b><i>Facebook</i></b> de les <b>Informacions personals</b>, en l'Opac Web; 
		<br />
		<br />
		1.5. Confirmeu les modificacions. 
		<br />
		<br />
		<div class='help-facebook-info-es-ca'></div>
	</div>
	<br />
	<br />
	<b>2. Afegiu l'usuari de la biblioteca com a amic de <i>Facebook</i>.</b>
	<br />
	<br />
	<div style="margin-left:10px;">
		<% if global_facebook_perfil <> "" then %>
			2.1. Accediu al link <a href="<%= global_facebook_perfil %>" target="_blank"><%= global_facebook_perfil %></a>;
		<% else %>
			2.1. Accediu al link;
		<% end if %>
		<br />
		<br />
		2.2. Entreu al compte de <i>Facebook</i>;
		<br />
		<br />
		2.3. Feu clic en el botó "+ Afegeix com a amistat":
		<br />
		<br />
		<div class='help-facebook-add-friend-es-ca'></div>
		<br />
		<br />
		2.4. Confirmació del sistema.
		<br />
		<br />
		Observacions: Fins que l'usuari de la biblioteca no s'hagi afegit com a amic de <i>Facebook</i>, els missatges de la biblioteca només es podran rebre en la safata d'entrada "Altres", a la qual es pot accedir segons s'indica en la finestra següent:
		<br />
		<br />
		<div class='help-facebook-other-inbox-es-ca'></div>
	</div>
</div>
</body>
</html>