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
	<b>Receiving Library Messages on <i>Facebook</i> - Instructions</b>
</p>
<div style="text-align: left;margin-left:10px;">
	<br />
	<br />
	Register a Facebook username on SophiA’s Opac Web and you can receive e-mails from the library on the desktop of your <i>Facebook</i> account.
	<br />
	To do this please follow the instructions below: 
	<br />
	<br />
	<b>1. Register a <i>Facebook</i> username in the <i>Personal Information</i> field on the Opac Web.</b>
	<br />
	<br />
	<div style="margin-left:10px;">
		1.1. Access the link <a href="https://www.facebook.com/settings?tab=account&section=username" target="_blank">https://www.facebook.com/settings?tab=account&amp;section=username</a>;
		<br />
		<br />
		1.2. Enter into your <i>Facebook</i> account;
		<br />
		<br />
		1.3. Copy the name in the <b>Username</b> field:
		<br />
		<br />
        <div class='help-facebook-username-en-us'></div>
		<br />
		<br />
		1.4. Paste the name in the <b><i>Facebook</i></b> field in the <b>Personal information</b> window on the Opac Web; 
		<br />
		<br />
		1.5. Confirm the changes.
		<br />
		<br />
        <div class='help-facebook-info-en-us'></div>
	</div>
	<br />
	<br />
	<b>2. Add the library user as a friend on <i>Facebook</i></b>
	<br />
	<br />
	<div style="margin-left:10px;">
		<% if global_facebook_perfil <> "" then %>
			2.1. Access the link <a href="<%= global_facebook_perfil %>" target="_blank"><%= global_facebook_perfil %></a>;
		<% else %>
			2.1. Access the library’s <i>Facebook</i> profile;
		<% end if %>
		<br />
		<br />
		2.2. Enter into your <i>Facebook</i> account;
		<br />
		<br />
		2.3. Click on the "+ Add Friend" tab:
		<br />
		<br />
		<div class='help-facebook-add-friend-en-us'></div>
		<br />
		<br />
		2.4. System confirmation.
		<br />
		<br />
		Observations: until the library user has been added as a friend on <i>Facebook</i> library messages will only be received in the "Other" incoming mail section, which you can access by following the indications in the window below:
		<br />
		<br />
		<div class='help-facebook-other-inbox-en-us'></div>
	</div>
</div>
</body>
</html>