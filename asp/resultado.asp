<!-- #include file="../libasp/funcoes.asp" -->
<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->
<%
if (sMsgErro <> "") then
	Response.Write sMsgErro
else%>
	<!-- #include file="../asp/ler_parametros_busca.asp" -->
	<%

	sURL_RESUMO = "&dados=" & Server.URLEncode(sDados) & "&objeto=" & iObjeto & _
				  "&contexto=" & iContexto & "&imagem=" & iImagem & _
				  "&campo1=" & Server.URLEncode(sCampo1) & "&campo2=" & Server.URLEncode(sCampo2) & _
				  "&campo3=" & Server.URLEncode(sCampo3) & "&campo4=" & Server.URLEncode(sCampo4) & _
				  "&campo5=" & Server.URLEncode(sCampo5) & "&campo6=" & Server.URLEncode(sCampo6) & _
				  "&campo7=" & Server.URLEncode(sCampo7) & "&campo8=" & Server.URLEncode(sCampo8) & _
				  "&campo_ordenacao=" & CodigoCampoOrdenacao & _
				  "&campo_ordem=" & OrdemPesquisa

	sURL_RESULTADOS = "&dados=" &Server.URLEncode(sDados) & "&objeto=" & iObjeto & _
					  "&contexto=" & iContexto & "&imagem=" & iImagem & _
					  "&campo1=" & Server.URLEncode(sCampo1) & "&campo2=" & Server.URLEncode(sCampo2) & _
					  "&campo3=" & Server.URLEncode(sCampo3) & "&campo4=" & Server.URLEncode(sCampo4) & _
					  "&campo5=" & Server.URLEncode(sCampo5) & "&campo6=" & Server.URLEncode(sCampo6) & _
					  "&campo7=" & Server.URLEncode(sCampo7) & "&campo8=" & Server.URLEncode(sCampo8) & _
					  "&tmp=" & sTabTmp & "&tmp_objeto=" & iTmpObjeto & "&tmp_contexto=" & iTmpContexto & "&pagina=" & iPagina & _
					  "&campo_ordenacao=" & CodigoCampoOrdenacao & _
					  "&campo_ordem=" & OrdemPesquisa

%>
<div id="divTituloPesquisa">
	<span style="float: left">
		<a class="link_menu" href="#" onClick=NovaPesquisa()>Home</a> > <a class="link_menu" href="#" onClick="LinkVoltar('detalhe_resultados','<%=sURL_RESUMO%>')">Resumo</a> > Resultado
	</span>
    <!-- #include file="../asp/botaoLogin.asp" -->
</div>
<br>
<%
	'INÍCIO DO DESTAQUE
	Set ParamPesq = ROServer.CreateComplexType("TPesquisa")
	ParamPesq.sPalavraChave = sDados
	ParamPesq.iMaterial     = iObjeto
	ParamPesq.iContexto 	= iContexto
	ParamPesq.iImagem   	= iImagem
	ParamPesq.sCampo1		= sCampo1
	ParamPesq.sCampo2		= sCampo2
	ParamPesq.sCampo3		= sCampo3
	ParamPesq.sCampo4		= sCampo4
	ParamPesq.sCampo5		= sCampo5
	ParamPesq.sCampo6		= sCampo6
	ParamPesq.sCampo7		= sCampo7
	ParamPesq.sCampo8		= sCampo8
	ParamPesq.CampoOrdenacao= CodigoCampoOrdenacao
	ParamPesq.OrdemPesquisa	= OrdemPesquisa

	'FIM DO DESTAQUE
	
	xmlResultado = ROService.GetResultado(sTabTmp,iTmpObjeto,iTmpContexto,ParamPesq,iPagina)
	sMsgErro = TrataErros
	
	if (sMsgErro <> "") then
		Response.Write sMsgErro
	else
		Response.Write FormataResultados(xmlResultado,sURL_RESULTADOS)
	end if
	
	Set ParamPesq = nothing
	
end if

Set ROService = nothing
Set ROServer = nothing
%>