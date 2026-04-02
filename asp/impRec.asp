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
<table style="width:644; background-color:#FFFFFF;">
<tr>
<td class="td_imp td_right_top">
<%

iIndexSrv = IntQueryString("servidor", 1)
'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

if(config_multi_servbib = 1)then
	sNomeBib = arNomeSrv(iIndexSrv)
else
	sNomeBib = global_instituicao
end if

Response.write "<table class='table_resultados' style='border-spacing: 1px; padding: 4px'>"
Response.write "<tr>"
Response.write "	<td class='td_resultados td_center_top'>"
Response.write "		<br /><div class='centro'><br /><b>"&sNomeBib&"</b><br /><br />"&getTermo(global_idioma, 589, "Recibo de renovação", 0)&"</div><br /><hr style='color: #666666'>"
Response.write "	</td>"
Response.write "</tr>"
Response.write "<tr>"
Response.write "	<td class='td_resultados td_left_top'>"
Response.write "        <div id='dRecibo' style='width: 100%;'></div>"
Response.write "	</td>"
Response.write "</tr>"
Response.write "<tr>"
Response.write "	<td class='td_resultados td_left_top'>&nbsp;</td>"
Response.write "</tr>"
Response.write "</table>"
Response.write "</td>"
Response.write "</tr>"
Response.write "<tr class='td_right_bottom'><td>"
Response.write "SophiA Biblioteca"
%>
</td>
</tr>
</table>
<script type='text/javascript'>
	if (parent != null) {
		if (parent.parent != null) {
			var divPai   = parent.parent.document.getElementById('div_recibo');
			var divLocal = document.getElementById('dRecibo');
			if ((divPai != null) && (divLocal != null)) {
				divLocal.innerHTML = divPai.innerHTML;
			}
		}
	}
</script>
</body>
</html>