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
	<title>::<%=getTermo(global_idioma, 8316, "Favoritos", 0)%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge" />
	<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />
	<link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />
<% 
if config_css_estatico = 1 then 
	Response.Write "<link href='../inc/estilo.css' rel='stylesheet' type='text/css' />"
else 
	Response.Write "<link href='../libasp/estilo.asp?idioma="&global_idioma&"' rel='stylesheet' type='text/css' />"
end if 
%>
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


	<script type="text/javascript" src="../scripts/salvar_favoritos.js"></script>
</head>

<%
sCodigos = request.QueryString("codigos")
CodigoLista = IntQueryString("codigoLista", -1)
iIndexSrv = IntQueryString("servidor", 1)

if (sCodigos = "") then
	sSessao = "mysel"&iIndexSrv
	arMySel = split(Session(sSessao),",")
	for i = 0 to ubound(arMySel)		
		arCodigo = split(arMySel(i),".")
		if (cods <> "") then
			cods = cods + ","
		end if
		cods = cods + arCodigo(0)
	next
	sCodigos = cods
end if

iIdioma = IntQueryString("iIdioma", 0)
'O índice iIndexSrv que define em qual servidor será acessado
%>
	<!-- #include file="../libasp/updChannelProperty.asp" -->
	<body class="popup" onload="parent.fechaLoadingPopup();">
    <div class="div_enviar_favoritos">

<%	if sCodigos <> "" then %>

		<table style="height: 84px;" class="remover_bordas_padding max_width">
			<tr>
				<td>&nbsp;&nbsp;<%=getTermo(global_idioma, 8319, "Selecione uma lista de favoritos:", 0)%></td>
			</tr>	
			<tr>
				<td >
					<div class="div_selecao_favoritos" id="div_selecao_favoritos">
						<!-- #include file="monta_lista_favoritos.asp" -->
						<div id="div_inputListaFavorito">
							<div class="div_inputListaFavorito"><%=getTermo(global_idioma, 8320, "Descrição da nova lista de favoritos", 0)%></div>
							<div><input  class="inputListaFavorito" type="text" name="inputListaFavorito" id="inputListaFavorito" value=''  maxlength="50" /></div>
						</div>
					</div>
				</td>
			</tr>
			<tr>
				<td>
					<div class="div_selecao_favoritos_botoes">
					<%	
						Response.Write "<input type='button' name='submit' value='"&getTermo(global_idioma, 4, "Confirmar", 0)&"' onclick=""incluirFavoritos('"&sCodigos&"',"&iIndexSrv&","&iIdioma&");"">"
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
						Response.Write "<input type='button' onclick=parent.fechaPopup('asp'); value='"&getTermo(global_idioma, 5, "Cancelar", 0)&"' />"
					 %>
					</div>
				</td>
			</tr>
	    </table>
	<% else 
		Response.Write("<script type='text/javascript'>alert('"&getTermo(global_idioma, 419, "Código inválido.", 0)&" "&getTermo(global_idioma, 3794, "Por favor, tente novamente.", 0)&"');</script>")
		Response.Write("<script type='text/javascript'>parent.fechaPopup();</script>")
	end if
%>
    </div>
	</body>
</html>
