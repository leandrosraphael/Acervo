<!-- #include file="libasp/funcoes.asp" -->
<%
TipoOperacao = Request.Form("Tipo_Operacao")

'Operação de LOGIN
if(TipoOperacao = "LOGIN")then
	'Pega o Login do usuário
	sLogin   = Request.Form("loginext")
	
	'Grava a sessão 
	if(sLogin <> "") then
		Session("loginext") = sLogin	
	end if
	
	'Redireciona para a página principal do Terminal Web
	Response.Redirect "index.html"

'Operação de LOGOUT	
elseif(TipoOperacao = "LOGOUT")then
	LiberarSessoes true
	
	sURL = Request.Form("URL_Redir")
	
	Response.Redirect sURL
else
	'Redireciona para a página principal do Terminal Web
	Response.Redirect "index.html"
end if
%>