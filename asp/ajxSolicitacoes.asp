<!DOCTYPE html>
<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
	<%
	sDiretorioArq = "asp"
	%>
	<!-- #include file="../config.asp" -->
	<!-- #include file="../idiomas/idiomas.asp" -->
	<!-- #include file="../libasp/header.asp" -->
	<!-- #include file="../libasp/funcoes.asp" -->
	<title></title>
	<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css"> 
    <link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />
    <% if config_css_estatico = 1 then %>
		<link href="../inc/estilo.css" rel="stylesheet" type="text/css"> 
    <% else %>
		<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css">
	<% end if %>
	<link href="../inc/contraste.css" rel="stylesheet" type="text/css">
    <link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
	</head>
<body>
<%
if config_multi_servbib = 1 then
	iIndexSrv = Session("Servidor_Logado")

	if iIndexSrv = "" then
		iIndexSrv = 1
	end if

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
end if

On Error Resume Next

codigos = Request.QueryString("codigos")
QtdePag = Request.QueryString("total")
pag_atual = Request.QueryString("pagina")
TotalSolic = Request.QueryString("solicitacoes")

Set ROService = ROServer.CreateService("Web_Consulta")
XMLHistorico = ROService.MontarHistoricoSolicitacoesEmprestimo(CStr(codigos),global_idioma)

HTML_HIST = ""

if (left(XMLHistorico,5) = "<?xml") then
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml XMLHistorico
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "solicitacoes" then
		HTML_HIST = HTML_HIST & "<p class='centro'><b>"&getTermo(global_idioma, 7340, "Histórico de solicitações de empréstimo", 0)&"</b></p><br/>"				
		'*************************************
		'Exibe paginação
		'*************************************
		if (CInt(QtdePag) > 1) then
			sNavegador = "<b>" & TotalSolic & "</b> "&getTermo(global_idioma, 7341, "solicitações encontradas", 2)&"&nbsp;&nbsp;-&nbsp;&nbsp;<b>" & QtdePag & "</b> "&getTermo(global_idioma, 236, "Páginas", 0)&"&nbsp;&nbsp;&nbsp;&nbsp;"
			if (CInt(QtdePag) > 5) then
		        sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1325, "Primeira", 0)&"' style='cursor:pointer;' onClick=LinkSolicitacoes("&TotalSolic&",1,"&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-first '></span></a>"					
                sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1326, "Anterior", 0)&"' style='cursor:pointer;' onClick=LinkSolicitacoesAnt("&TotalSolic&","&pag_atual&","&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-previous '></span></a>"				
				if (CInt(pag_atual) <= 3) then
					iIDPag = 1
				elseif ((CInt(QtdePag) - CInt(pag_atual)) >= 2) then
					iIDPag = CInt(pag_atual) - 2
				else 
					iIDPag = CInt(QtdePag) - 4
				end if
			else
				iIDPag = 1
			end if
			iCountPag = 1
			do while ((iCountPag <= 5) and (iIDPag <= CInt(QtdePag)))
				sLink = "<a class='link_pag' title='"&getTermo(global_idioma, 1323, "Página", 0)&" "&iIDPag&"' href=""javascript:LinkSolicitacoes("&TotalSolic&","&iIDPag&","&QtdePag&");"">"
				if (CInt(pag_atual) = iIDPag) then
					sLink = sLink & "<b><span class='paginacao_link_atual'>"
				end if
				sLink = sLink & iIDPag
				if (CInt(pag_atual) = iIDPag) then
					sLink = sLink & "</span></b>"
				end if
				sNavegador = sNavegador & sLink & "</a>&nbsp;"
				iIDPag = iIDPag + 1
				iCountPag = iCountPag + 1
			loop
			if (QtdePag > 5) then
		        sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1327, "Próxima", 0)&"' style='cursor:pointer;' onClick=LinkSolicitacoesProx("&TotalSolic&","&pag_atual&","&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-next '></span></a>"					
                sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1328, "Última", 0)&"' style='cursor:pointer;' onClick=LinkSolicitacoes("&TotalSolic&","&QtdePag&","&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-last '></span></a>"							
			end if
			HTML_HIST = HTML_HIST & "<table style=' width: 100%;'>"
			HTML_HIST = HTML_HIST & "<tr style=' height: 25px;'><td class='td_servicos_titulo td_left_middle'>&nbsp;"&sNavegador&"&nbsp;</td></tr></table>"
		end if
					
		HTML_HIST = HTML_HIST & "<table class='tab_circulacoes max_width' style=' border-spacing: 1px;'>"

		Registro = 1						

		For Each xmlSolic In xmlRoot.childNodes
		
			sCabecalho = ""
			sRegistro  = ""
						
			if (Registro mod 2) > 0 then '### IMPAR
				fontcolor = "#000000" 	
				td_class = "td_tabelas_valor2"
				link_class = "link_serv"
			else '### PAR
				fontcolor= "#000000" 
				td_class = "td_tabelas_valor1"
				link_class = "link_serv"								
			end if				
						
			if xmlSolic.nodeName  = "solicitacao" then
				For Each xmlCampos In xmlSolic.childNodes
					'*************************************
					'SEQUENCIAL
					'*************************************
					if xmlCampos.nodeName  = "sequencial" then
						sequencial_desc = xmlCampos.attributes.getNamedItem("descricao").value
						sequencial = xmlCampos.attributes.getNamedItem("valor").value
						'Monta Registro
						sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&sequencial&"&nbsp;</td>"
						'Monta Cabeçalho
						if Registro = 1 then
							sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 25px'>&nbsp;"&sequencial_desc&"&nbsp;</td>"
						end if
					end if
					'*************************************
					'TITULO
					'*************************************
					if xmlCampos.nodeName  = "titulo" then
						titulo_desc = xmlCampos.attributes.getNamedItem("descricao").value
						titulo = xmlCampos.attributes.getNamedItem("valor").value
						'Monta Registro
						sRegistro = sRegistro & "<td class='"&td_class&" esquerda'>&nbsp;<span style='color: "&fontcolor&"'>"&titulo&"</span><br />"
						'Monta Cabeçalho
						if Registro = 1 then
							sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro'>&nbsp;"&titulo_desc&"&nbsp;</td>"
						end if
					end if
					'*************************************
					'MATERIAL
					'*************************************
					if xmlCampos.nodeName  = "material" then
						material_desc = xmlCampos.attributes.getNamedItem("descricao").value
						material = xmlCampos.attributes.getNamedItem("valor").value
						'Monta Registro
						sRegistro = sRegistro & "&nbsp;<span style='color:#999999;font-weight:600'>" & material_desc & ":</span>&nbsp;" & material
					end if
					'*************************************
					'BIBLIOTECA
					'*************************************
					if xmlCampos.nodeName  = "biblioteca" then
						biblioteca_desc = xmlCampos.attributes.getNamedItem("descricao").value
						biblioteca = xmlCampos.attributes.getNamedItem("valor").value
                        if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB_EXEMPLAR").value = 1) then
                            cod_bib = xmlCampos.attributes.getNamedItem("CODIGO_BIB_EXEMPLAR").value
                            descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(cod_bib) & "," & veio_de_popup & ",0);'>"&biblioteca&"</a>"
                        else
                            descricaoBib = biblioteca
                        end if
						'Monta Registro
						sRegistro = sRegistro & "&nbsp;<span style='color:#999999;font-weight:600'>" & biblioteca_desc & ":</span>&nbsp;" & descricaoBib & "</td>"
					end if
					'*************************************
					'TOMBO / CODIGO
					'*************************************
					if xmlCampos.nodeName  = "tombo" then
						tombo_desc = xmlCampos.attributes.getNamedItem("descricao").value
						tombo = xmlCampos.attributes.getNamedItem("valor").value
						'Monta Registro
						sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&tombo&"&nbsp;</td>"
						'Monta Cabeçalho
						if Registro = 1 then
							sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 100px'>&nbsp;"&tombo_desc&"&nbsp;</td>"
						end if
					end if
					'*************************************
					'PROTOCOLO
					'*************************************
					if xmlCampos.nodeName  = "protocolo" then
						protocolo_desc = xmlCampos.attributes.getNamedItem("descricao").value
						protocolo = xmlCampos.attributes.getNamedItem("valor").value
						'Monta Registro
						sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&protocolo&"&nbsp;</td>"
						'Monta Cabeçalho
						if Registro = 1 then
							sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 80px'>&nbsp;"&protocolo_desc&"&nbsp;</td>"
						end if
					end if
					'*************************************
					'DATA SOLICITAÇÃO
					'*************************************
					if xmlCampos.nodeName  = "data_solic" then
						data_solic_desc = xmlCampos.attributes.getNamedItem("descricao").value
						data_solic = xmlCampos.attributes.getNamedItem("valor").value
						'Monta Registro
						sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&data_solic&"&nbsp;</td>"
						'Monta Cabeçalho
						if Registro = 1 then
							sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 80px'>&nbsp;"&data_solic_desc&"&nbsp;</td>"
						end if
					end if
					'*************************************
					'BIBLIOTECA RETIRADA
					'*************************************
					if xmlCampos.nodeName  = "biblioteca_retirada" then
						biblioteca_retirada_desc = xmlCampos.attributes.getNamedItem("descricao").value
						biblioteca_retirada = xmlCampos.attributes.getNamedItem("valor").value
                        if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB_DESTINO").value = 1) then
                            cod_bib_retirada = xmlCampos.attributes.getNamedItem("CODIGO_BIB_DESTINO").value
                            descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(cod_bib_retirada) & "," & veio_de_popup & ",0);'>"&biblioteca_retirada&"</a>"
                        else
                            descricaoBib = biblioteca_retirada
                        end if
						sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&descricaoBib&"&nbsp;</td>"
						'Monta Cabeçalho
						if Registro = 1 then
							sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 120px'>&nbsp;"&biblioteca_retirada_desc&"&nbsp;</td>"
						end if
					end if
					'*************************************
					'DATA VENCIMENTO
					'*************************************
					if xmlCampos.nodeName  = "vencimento" then
						vencimento_desc = xmlCampos.attributes.getNamedItem("descricao").value
						vencimento = xmlCampos.attributes.getNamedItem("valor").value
						'Monta Registro
						sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&vencimento&"&nbsp;</td>"
						'Monta Cabeçalho
						if Registro = 1 then
							sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 80px'>&nbsp;"&vencimento_desc&"&nbsp;</td>"
						end if
					end if
					'*************************************
					'SITUAÇÃO
					'*************************************
					if xmlCampos.nodeName  = "situacao" then
						situacao_desc = xmlCampos.attributes.getNamedItem("descricao").value
						situacao = xmlCampos.attributes.getNamedItem("valor").value
						'Monta Registro
						sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&situacao&"&nbsp;</td>"
						'Monta Cabeçalho
						if Registro = 1 then
							sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 100px'>&nbsp;"&situacao_desc&"&nbsp;</td>"
						end if
					end if
				Next
							
				if Registro = 1 then
					HTML_HIST = HTML_HIST & "<tr style=' height: 20px;'>"
					HTML_HIST = HTML_HIST & sCabecalho
					HTML_HIST = HTML_HIST & "</tr>"
				end if
				HTML_HIST = HTML_HIST & "<tr style=' height: 20px;'>" & sRegistro & "</tr>"
			end if
						
			Registro = Registro + 1
		Next						
		HTML_HIST = HTML_HIST & "</table><br />"						
	End if	
	Set xmlDoc = nothing
	Set xmlRoot = nothing	
End if

Set ROService = nothing
		
Response.Write "<script type='text/javascript'>"
Response.Write "	var HTML_HIST = """ & replace(HTML_HIST,"""","\""") & """;"
Response.Write "	var DIV = parent.document.getElementById('div_solic');"
Response.Write "	if (DIV != null) {"
Response.Write "		DIV.innerHTML = HTML_HIST;"
Response.Write "	}"
Response.Write "	var DIV = parent.document.getElementById('solic_loading');"
Response.Write "	if (DIV != null) {"
Response.Write "		DIV.innerHTML = '&nbsp;';"
Response.Write "	}"
Response.Write "</script>"

%>
</body>
</html>