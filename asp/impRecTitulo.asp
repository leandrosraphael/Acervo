<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
	<title></title>
<% sDiretorioArq="asp" %>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link rel="stylesheet" href="../inc/estilo_imp.css" />
<link rel="stylesheet" href="../inc/contraste.css" />

</head>
<body onload="parent.parent.fechaLoadingPopup();">
<%

iIndexSrv = IntQueryString("servidor", 1)
'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

codigo_obra = Request.QueryString("codigo")

Set ROService = ROServer.CreateService("Web_Consulta")	
ReciboTitulo = ROService.MontaReciboFapcom(codigo_obra)
Set ROService = nothing

TrataErros(1)

Response.write "<table class='table_resultados' style='border-spacing: 1px; padding: 4px' >"
Response.write "<tr>"
Response.write "	<td class='td_resultados td_left_top'>"
Response.write "	<span style='font: 10px arial'>"
Response.write replace(ReciboTitulo, Chr(10), "<br/>")
Response.write "	</span>"
Response.write "	</td>"
Response.write "</tr>"
Response.write "</table>"
%>
</body>
</html>