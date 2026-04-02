<% 
sDiretorioArq="asp" 
Response.ContentType = "text/xml"
nao_imprime_variaveis_globais = "1"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%
	iIndexSrv = IntQueryString("servidor", 1)

	if iIndexSrv = "" then
		iIndexSrv = 1
	end if

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	
	usuario = Request.QueryString("usuario")
	obra = Request.QueryString("codigo")
	ano = Request.QueryString("ano")
	volume = Request.QueryString("volume")
	numero = Request.QueryString("numero")
	
	if Request.QueryString("suporte") = "" then
		suporte = 0
	else
		suporte = Request.QueryString("suporte")
	end if
	
	if Request.QueryString("biblioteca") = "" then
		biblioteca = 0
	else
		biblioteca = Request.QueryString("biblioteca")
	end if
	
	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	xml_dados_opc = ROService.DadosOpcionais(usuario,obra,CStr(ano),CStr(volume),CStr(numero),CInt(suporte),CInt(biblioteca))
	TrataErros(1)
	SET ROService = nothing
	
	' Para resolver conflito de codificação, deixando os acentos de forma correta
    Response.Write replace(xml_dados_opc, "windows-1252", "utf-8")
%>