	<%
sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%
	iIndexSrv = 1
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	
	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	SET Retorno = ROServer.CreateComplexType("TMensagem")
	SET DadosUsuarioExterno = ROServer.CreateComplexType("TDadosUsuarioExterno")

	origem_mobile = Request.Form("origem_mobile")
	DadosUsuarioExterno.nome = Request.Form("nome")
	DadosUsuarioExterno.email = Request.Form("email")
	DadosUsuarioExterno.senha = Request.Form("senha")
	DadosUsuarioExterno.confirmacao_senha = Request.Form("confirmacao_senha")
	DadosUsuarioExterno.codigo_empresa = Request.Form("codigo_empresa")
	DadosUsuarioExterno.nome_empresa = Request.Form("nome_empresa")
	DadosUsuarioExterno.termo_aceite = Request.Form("termo_aceite")
    global_idioma = Request.QueryString("idioma")
	DadosUsuarioExterno.idioma = global_idioma

	SET Retorno = ROService.CadastrarUsuarioExterno(DadosUsuarioExterno)

	TrataErros(1)
	
	SET ROService = nothing
	SET DadosUsuarioExterno = nothing
	if (origem_mobile = "1") then
	    response.clear
		if (Retorno.Result) then
			response.write "1|" & Retorno.sMsg
		else
			response.write "0|" & Retorno.sMsg
		end if
		response.end
	else %>
			<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />
			<% if config_css_estatico = 1 then %>
				<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
			<% else %>
				<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
			<% end if %>
			<div class="centro div-popup">
				<p class="titulo-link-popup mensagem-popup"><%= Retorno.sMsg %></p>
			</div>
			<div class="centro" style="margin-bottom: 20%;">
				<% if (Retorno.Result) then %>
					<input type="button" onClick="parent.fechaPopup('asp');" value="<%=getTermo(global_idioma, 220, "Fechar", 0)%>"
				<% else %>
					<input type="button" onClick="javascript: retornarCadastro();" value="<%=getTermo(global_idioma, 1386, "voltar", 0)%>"
				<% end if %>
			</div>
			<% SET Retorno = nothing %>
			<script type="text/javascript">
				function retornarCadastro() {
					$("div#mensagem").hide();
					$("div#conteudo-principal").show();
				}
			</script>
	<%
	end if
	%>