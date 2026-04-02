<%
	sDiretorioArq = "asp"
    nao_imprime_variaveis_globais = "1"
%>

<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<%
if config_multi_servbib = 1 then
	iIndexSrv = Session("Servidor_Logado")

	if iIndexSrv = "" then
		iIndexSrv = 1
	end if

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
end if
%>
<body>
<script type="text/javascript">
	function VoltarTelaDevolucao(modo) {
		parent.mainFrame.location = "../index.asp?modo_busca=" + modo + "&content=circulacoes&acao=emprestimo" + getGlobalUrlParams();
	}
</script>
<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top" style="display: block">
<%
On Error Resume Next

	Set ROService = ROServer.CreateService("Web_Consulta")
	Set DadosEmprestimo = ROServer.CreateComplexType("TDadosEmprestimo")
	
	DadosEmprestimo.Biblioteca = Request.QueryString("biblioteca")
	DadosEmprestimo.Usuario = Request.QueryString("usuario")
	DadosEmprestimo.CodigoBarras = Request.QueryString("codigobarras")
	DadosEmprestimo.UsuarioTerminalWeb = Request.QueryString("usuariotw")
	DadosEmprestimo.Material = Request.QueryString("material")
	DadosEmprestimo.ReservaOrigem = Request.QueryString("reserva")
	DadosEmprestimo.SolicitacaoEmprestimo = Request.QueryString("solicitacaoemprestimo")
	DadosEmprestimo.Registro = Request.QueryString("registro")
	DadosEmprestimo.Tombo = Request.QueryString("tombo")
	DadosEmprestimo.TipoEmprestimo = Request.QueryString("tipoemprestimo")

	sXMLRenova = ROService.EmprestarExemplar(DadosEmprestimo)
	Set ROService = nothing
	Set DadosEmprestimo = nothing
	TrataErros(1)

	if (left(sXMLRenova,5) = "<?xml") then
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml sXMLRenova
		Set xmlRoot = xmlDoc.documentElement
		if xmlRoot.nodeName = "EMPRESTIMO" then
			For Each xmlPNode In xmlRoot.childNodes
				if xmlPNode.nodeName = "INFOUSU" then
					sGridRenovacao = sGridRenovacao & "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
					sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'><td class='td_tabelas_titulo centro' colspan=2><b>"&getTermo(global_idioma, 9009, "Dados da empréstimo", 0)&"</b><br /></td></tr>"
					For Each xmlUsu In xmlPNode.childNodes
						if xmlUsu.nodeName  = "NOME_USU" then	
							Nome_Usuario = xmlUsu.attributes.getNamedItem("VALOR").value
							Desc_Nome_Usuario = xmlUsu.attributes.getNamedItem("DESCRICAO").value
							sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
							sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width:172px'>&nbsp;"&Desc_Nome_Usuario&"</td>"
							sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&Nome_Usuario&"</td>"
							sGridRenovacao = sGridRenovacao & "</tr>" 
						end if
						if xmlUsu.nodeName  = "CODIGO_USU" then	
							RM_Usuario = xmlUsu.attributes.getNamedItem("VALOR").value
							Desc_RM_Usuario = xmlUsu.attributes.getNamedItem("DESCRICAO").value
							sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
							sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width: 172px'>&nbsp;"&Desc_RM_Usuario&"</td>"
							sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&RM_Usuario&"</td>"
							sGridRenovacao = sGridRenovacao & "</tr>" 
						end if
					Next
					sGridRenovacao = sGridRenovacao & "</table>"
				end if
				if xmlPNode.nodeName = "INFO_EMPRESTIMO" then
					sGridRenovacao = sGridRenovacao & "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
					For Each xmlEmp In xmlPNode.childNodes
						sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
						sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width:172px'>&nbsp;"&xmlEmp.attributes.getNamedItem("DESCRICAO").value&"</td>"
						sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&xmlEmp.attributes.getNamedItem("VALOR").value&"</td>"
						sGridRenovacao = sGridRenovacao & "</tr>" 
					Next
					sGridRenovacao = sGridRenovacao & "</table>"
				end if
				if xmlPNode.nodeName = "MENSAGEM_EMPRESTIMO" then
					For Each xmlMsg In xmlPNode.childNodes
						if xmlMsg.nodeName = "MENSAGEM" then
							sGridRenovacao = sGridRenovacao & "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
							sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'><td class='td_tabelas_titulo centro' colspan=2><b>"&getTermo(global_idioma, 725, "Observações", 0)&"</b><br /></td></tr>"
							sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
							sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width:172px'>&nbsp;"&xmlMsg.attributes.getNamedItem("DESCRICAO").value&"</td>"
							sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&xmlMsg.attributes.getNamedItem("VALOR").value&"</td>"
							sGridRenovacao = sGridRenovacao & "</tr>" 
							sGridRenovacao = sGridRenovacao & "</table>"
						end if
					Next
				end if
			Next
		end if
		
		Response.Write "<br />"
		Response.write "<table style='width: 100%'>"
		Response.write "<tr style='height: 25px'>"
		Response.write "<td class='td_servicos_titulo td_left_middle'>"
		Response.write "</td></tr></table>"
		Response.Write sGridRenovacao
		Response.Write "<br />"
	end if
%>
</body>