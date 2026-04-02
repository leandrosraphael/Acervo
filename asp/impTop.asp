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
<link rel="stylesheet" href="../inc/estilo_padrao.css" />
<link rel="stylesheet" href="../inc/estilo_imp.css" />
<link rel="stylesheet" href="../inc/imagem_padrao.css" />
<link rel="stylesheet" href="../inc/contraste.css" />
<link rel="stylesheet" href="../inc/imagem_contraste.css" />

<%
codigoFavorito = IntQueryString("codigoFavorito", 0) 
iIndexSrv = IntQueryString("Servidor", 1)
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

On Error Resume Next
SET ROService = ROServer.CreateService("Web_Consulta")
Set objCfgSolic = ROServer.CreateComplexType("TSolicConsulta")
Set objCfgSolic = ROService.SolicConsulta

TrataErros(1)

if objCfgSolic.Solic_Impressao > 0 then
	bSolic = (global_IP_Local = 1)
else
	bSolic = false
end if

Set objCfgSolic = nothing
SET ROService = nothing

%>
<script type='text/javascript'>
// (C) 2004 www.CodeLifter.com
// Free for all users, but leave in this header
function framePrint(whichFrame){
	parent[whichFrame].focus();
	parent[whichFrame].print();
}

function LinkRefbib(servidor, codigoFavorito) {
	parent.parent.exibeLoadingPopup();
	parent.imp_impFrame.location = 'refbib.asp?Servidor='+servidor+'&codigoFavorito='+codigoFavorito+'&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>';
	<% if (bSolic = true) then %>
	exibeSolic();
	<% end if %>
}

function LinkMinhasSels(servidor, codigoFavorito) {
	parent.parent.exibeLoadingPopup();
	parent.imp_impFrame.location = 'impSels.asp?Servidor='+servidor+'&codigoFavorito='+codigoFavorito+'&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>';
	<% if (bSolic = true) then %>
	escondeSolic();
	<% end if %>
}

function LinkSolicConsulta(servidor, codigoFavorito) {
	var pag = String(parent.imp_impFrame.location);
	if (pag.indexOf('impSels') < 0) {
		parent.parent.parent.hiddenFrame.popup_html = parent.parent.adicionarURL(parent.parent.parent.hiddenFrame.popup_html,'&Tipo=RB');
		parent.parent.abrePopup2('asp/lista_imp.asp?acao=conferir&Servidor='+servidor+'&codigoFavorito='+codigoFavorito+'&Tipo=RB&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>','<%=getTermo(global_idioma, 831, "Solicitar impressão", 0)%>',320,210,false,false);
	}
}

function escondeSolic() {
	var div = document.getElementById('div_solic');
	if (div != null) {
		div.style.display='none';
	}
}

function exibeSolic() {
	var div = document.getElementById('div_solic');
	if (div != null) {
		div.style.display='block';
	}
}

</script>
</head>

<body>
<% if (bSolic = true) then %>
	<table style="width: 615px; height: 40px" class="remover_bordas_padding">
<% else %>
	<table style="width: 445px; height: 40px" class="remover_bordas_padding">
<% end if %>
  <tr style="vertical-align: middle">
    <td class="centro" style="width: 130px; height: 20px"><a href="#" class="link_serv" onClick="LinkMinhasSels(<%=iIndexSrv%>,<%=codigoFavorito%>);"><span class="transparent-icon span_imagem icon_16 icon-small-document-b-h"></span>&nbsp;<%=getTermo(global_idioma, 1485, "Lista de materiais", 0)%></a></td>
    <td class="centro" style="width: 200px"><a href="#" class="link_serv" onClick="LinkRefbib(<%=iIndexSrv%>,<%=codigoFavorito%>);"><span class="transparent-icon span_imagem icon_16 icon-small-book-b-h"></span>&nbsp;<%=getTermo(global_idioma, 792, "Referência bibliográfica", 0)%></a></td>
	<td class="centro" style="width: 5px"><span class="span_imagem icon_1_18 barra-png"></span></td>
	<% if (bSolic = true) then %>
	    <td class="centro" style="width:130px"><a href="javascript:framePrint('imp_impFrame');" class="link_serv"><span class="transparent-icon span_imagem icon_16 icon-small-print-b-h"></span>&nbsp;<%=getTermo(global_idioma, 13, "Imprimir", 0)%></a></td>
	    <td class="centro" style="width:150px"><div id="div_solic"><a href="#" class="link_serv" onClick="LinkSolicConsulta(<%=iIndexSrv%>,<%=codigoFavorito%>);"><span class="transparent-icon span_imagem icon_16 icon-small-print-b-h"></span>&nbsp;<%=getTermo(global_idioma, 831, "Solicitar impressão", 0)%></a></div></td>
	<% else %>
    	<td class="centro" style="width:120px"><a href="javascript:framePrint('imp_impFrame');" class="link_serv"><span class="transparent-icon span_imagem icon_16 icon-small-print-b-h"></span>&nbsp;<%=getTermo(global_idioma, 13, "Imprimir", 0)%></a></td>
	<% end if %>
  </tr>
</table>
<% if (bSolic = true) then %>
	<script> 
		<% if (Request.QueryString("Tipo") = "RB") then %>
			exibeSolic();
		<% else %>
			escondeSolic();
		<% end if %>
	</script>
<% end if %>
</body>
</html>