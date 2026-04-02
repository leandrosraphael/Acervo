<%
sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<%

Tipo = request.QueryString("tipo")
Codigo = request.QueryString("codigo")

iIndexSrv = IntQueryString("Servidor", 1)
'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

On Error Resume Next
SET ROService = ROServer.CreateService("Web_Consulta")
Str_RefBib = ROService.MontaRefBib(Tipo,Codigo,0,0,global_idioma, true,true)

TrataErros(1)

SET ROService = nothing


str_final = "<table style= 'width: 100%; height: 70px; background-color: #CCCCCC; border-spacing: 1px; display: inline-table;'>"
str_final = str_final & "<tr><td style='background-color: #FFFFFF; height: 50px'><p class='centro' style='padding-top: 20px'>"
str_final = str_final & "<div class='justificado' style='padding: 7px;'>"
str_final = str_final & replace(replace(replace(replace(Str_RefBib,chr(10),"<br />"),chr(13), ""),"\","\\"),"""","\""")
str_final = str_final & "</div>"
str_final = str_final & "</p>"
str_final = str_final & "<br/><div class='direita' style='padding: 4px;'><span class='transparent-icon span_imagem icon_16 icon-small-delete-b-h'></span>&nbsp;"
str_final = str_final & "<a class='link_serv' onClick='tiraRefbib()' style='cursor:pointer;'>"&getTermo(global_idioma, 220, "fechar", 2)&"</div></td></tr></table><br/>"

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<script>
<!--
    window.onload = function () {
        var str_final = "<%= str_final %>";
        //alert(str_final)
        var console = parent.document.getElementById("div_refbib");
        console.innerHTML = str_final;
        //alert('aiaiaia');
    }
-->
</script>
<title></title>
</head>
<body></body>
</html>