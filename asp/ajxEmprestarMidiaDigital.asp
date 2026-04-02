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

Set TMsg = ROServer.CreateComplexType("TMensagem")
Set TMsg = ROService.EmprestarLivroDigitalFGV(Midia, Usuario, global_idioma)

sMsg = TMsg.sMsg
bResultado = TMsg.Result

Set TMsg = nothing
Set ROService = nothing

'******************************************
'Contagem de acesso
'******************************************
Set ROService = ROServer.CreateService("Web_Consulta")
ROService.ContaAcessosMidias 0, CLng(Midia), 1, CLng(Usuario), 0
Set ROService = nothing

TrataErros(1)
response.write Server.HTMLEncode(sMsg)

if (bResultado) then
	Response.Status = "200 OK"
else
	Response.Status = "203 Non-Authoritative Information"
end if

%>