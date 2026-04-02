<%
	'Desabilitando cache pois estava influenciando na passagem de parâmetros no site
	'Ao recarregar a mesma página os dados nem sempre eram atualizados
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>

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

%><!-- #include file="../libasp/updChannelProperty.asp" --><%	

	codMidia = Request.QueryString("midia")
    tipo = Request.QueryString("tipo")
    if len(trim(Session("codigo_usuario"))) = 0 then
	    usuario = 0
    else
	    usuario = Session("codigo_usuario")
    end if

	'******************************************
	'Contagem de acesso
	'******************************************
	Set ROService = ROServer.CreateService("Web_Consulta")
	ROService.ContaAcessosMidiaExemplar CLng(codMidia), CLng(tipo), CLng(usuario)
	Set ROService = nothing
	
%>