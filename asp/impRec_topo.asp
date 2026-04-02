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
<title>::<%=getTermo(global_idioma, 990, "Impressão", 0)%> - <%=getTermo(global_idioma, 589, "Recibo de renovação", 0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link rel="stylesheet" href="../inc/estilo_padrao.css" />
<link rel="stylesheet" href="../inc/estilo_imp.css" />
<link rel="stylesheet" href="../inc/imagem_padrao.css" />
<link rel="stylesheet" href="../inc/contraste.css" />
<link rel="stylesheet" href="../inc/imagem_contraste.css" />

<script type='text/javascript'>
// (C) 2004 www.CodeLifter.com
// Free for all users, but leave in this header
function framePrint(whichFrame){
	parent[whichFrame].focus();
	parent[whichFrame].print();
}

</script>
</head>

<body>
<table class="remover_bordas_padding">
  <tr style="vertical-align: middle">
   	<td class="centro" style="width: 100px; height: 40px"><a href="javascript:framePrint('imp_impFrame');" class="link_serv">
    	<span class="transparent-icon span_imagem icon_16 icon-small-print-b-h"></span>&nbsp;<%=getTermo(global_idioma, 13, "Imprimir", 0)%></a>
    </td>
  </tr>
</table>
</body>
</html>