<%
' Exclui Solicitação Empréstimo
codigo_solicitacao = Request.QueryString("codigo_solic")

if config_multi_servbib = 1 then
	iIndexSrv = Session("Servidor_Logado")

	if iIndexSrv = "" then
		iIndexSrv = 1
	end if

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
end if

SET ROService = ROServer.CreateService("Web_Consulta")
sMsg = ROService.ExcluirSolicitacaoEmprestimo(codigo_solicitacao, global_idioma)
SET ROService = nothing

TrataErros(1)

Response.Write "<script type='text/javascript'>alert('"&sMsg&"');</script>"

Response.Write "<script type='text/javascript'>parent.mainFrame.location='index.asp?modo_busca='+parent.hiddenFrame.modo_busca+'&content=solicitacoes_emp&iBanner='+parent.hiddenFrame.iBanner + '&iIdioma='+parent.hiddenFrame.iIdioma;</script>"
%>