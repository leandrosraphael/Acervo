<%
	sDiretorioArq = "asp"
	nao_imprime_variaveis_globais = "1"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<%

Midia = request.QueryString("midia")
Usuario = request.QueryString("usuario")

iIndexSrv = IntQueryString("Servidor", 1)
'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

On Error Resume Next
Set ROService = ROServer.CreateService("Web_Consulta")
sMsg = ROService.IncluirReservaDigital(Midia, Usuario, global_idioma)
Set ROService = nothing

TrataErros(1)

sMsg = replace(sMsg,"[QUEBRA]","")
sMsg = replace(sMsg,"[/QUEBRA]","")
sMsg = replace(sMsg,"[NEGRITO]","")
sMsg = replace(sMsg,"[/NEGRITO]","")

response.write Server.HTMLEncode(sMsg)

response.End
%>