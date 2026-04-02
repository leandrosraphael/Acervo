<%
	sDiretorioArq = "asp"
    nao_imprime_variaveis_globais = "1"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%
iIndexSrv = IntQueryString("iIndexSrv", 1)
iIdioma = IntQueryString("iIdioma", 0)
veio_de = "servicos"
listaSelecionada = IntQueryString("listaSelecionada", 0)

'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%>
<!-- #include file="../libasp/updChannelProperty.asp" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
</head>
<body>	
	<!-- #include file="monta_lista_favoritos.asp" -->
</body>
</html>
