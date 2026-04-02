<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top" style="display: block">
<%
	if Session("codigo_usuario") = "" then
		Response.write "<script type='text/javascript'>parent.mainFrame.location='asp/logout.asp?logout=sim&modo_busca=" & GetModo_Busca & "&iBanner="&global_tipo_banner&"&iIdioma=" & global_idioma & "';</script>"
	end if

	if Session("Logado") = "sim" then
	
		if config_multi_servbib = 1 then
			iIndexSrv = Session("Servidor_Logado")
	
			if iIndexSrv = "" then
				iIndexSrv = 1
			end if
	
			'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
			%><!-- #include file="../libasp/updChannelProperty.asp" --><%
		end if	
	
		codigo_usu = Session("codigo_usuario")

		'Apresenta detalhe da sugestão
		if Request.querystring("acao") = "detalhe" OR Request.querystring("codsgt") <> "" then
%> 			<!-- #include file="detalhe_sgt.asp" --> 
<%
		'Apresenta histórico com sugestões
		else
			'Cria grid com sugestões		
			Response.write "<br /><p class='centro'><b>"&getTermoHtml(global_idioma, 542, "Sugestões de aquisição", 0)&"</b></p><br /><br />"
			
			Response.write "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
			Response.write "<tr style='height: 25px'><td class='td_servicos_titulo direita'>"
			Response.write "<a class='link_serv' href=""javascript:LinkSugestao(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-lamp'></span>"
			Response.write "&nbsp;"&getTermo(global_idioma, 1336, "Nova sugestão", 0)&"&nbsp;&nbsp;</a></td></tr>"
			Response.write "</table>"

			'Busca todas as sugestões do usuário
			Set ROService = ROServer.CreateService("Web_Consulta")	
			sXML = ROService.HistoricoSugestoes(Session("codigo_usuario"),global_idioma)
			Set ROService = nothing
			TrataErros(1)

			'Monta grid de sugestões
			if (left(sXML,5) = "<?xml") then
				'Processa o XML com as sugestões		
				Set xmlDoc = CreateObject("Microsoft.xmldom")
				xmlDoc.async = False
				xmlDoc.loadxml sXML
				Set xmlRoot = xmlDoc.documentElement

				'Verifica se usuário possui sugestões
				if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
					iSeq = 0					
					sColunas = ""
								
					'Carrega cada uma das sugestões
					For Each xmlSugestao In xmlRoot.childNodes
						if xmlSugestao.nodeName  = "SUGESTAO" then
							iSeq = iSeq + 1
							
							sSeq = ""
							sData = ""
							sTitulo = ""
							sEdicao = ""
							sAutor = ""
							sEditora = ""
							sQuantidade = ""
							sSituacao = ""
							sPrioridade = ""
							
							For Each xmlDados In xmlSugestao.childNodes
								'Sequencial
								if xmlDados.nodeName  = "SEQUENCIAL" then
									sSeq = xmlDados.attributes.getNamedItem("VALOR").value
									if iSeq = 1 then
										sDescCol = xmlDados.attributes.getNamedItem("DESCRICAO").value
										sColunas = sColunas & "<td class='td_tabelas_titulo centro' style='width: 25px'>"&sDescCol&"</td>"
									end if
								end if
								
								'Data
								if xmlDados.nodeName  = "DATA" then
									sData = xmlDados.attributes.getNamedItem("VALOR").value
									if iSeq = 1 then
										sDescCol = xmlDados.attributes.getNamedItem("DESCRICAO").value
										sColunas = sColunas & "<td class='td_tabelas_titulo centro' style='width: 80px'>"&sDescCol&"</td>"
									end if
								end if
							
								'Titulo
								if xmlDados.nodeName  = "TITULO" then
									sTitulo =  xmlDados.attributes.getNamedItem("VALOR").value
									if iSeq = 1 then
										sDescCol = xmlDados.attributes.getNamedItem("DESCRICAO").value
										sColunas = sColunas & "<td class='td_tabelas_titulo centro' style='width: 180px'>"&sDescCol&"</td>"
									end if
								end if
		
								'Edição
								if xmlDados.nodeName  = "EDICAO" then
									sEdicao = xmlDados.attributes.getNamedItem("VALOR").value
									if iSeq = 1 then
										sDescCol = xmlDados.attributes.getNamedItem("DESCRICAO").value
										sColunas = sColunas & "<td class='td_tabelas_titulo centro' style='white-space: nowrap'>"&sDescCol&"</td>"
									end if
								end if
		
								'Autor
								if xmlDados.nodeName  = "AUTOR" then
									sAutor = xmlDados.attributes.getNamedItem("VALOR").value
									if iSeq = 1 then
										sDescCol = xmlDados.attributes.getNamedItem("DESCRICAO").value
										sColunas = sColunas & "<td class='td_tabelas_titulo centro'>"&sDescCol&"</td>"
									end if
								end if
								
								'Editora
								if xmlDados.nodeName  = "EDITORA" then
									sEditora = xmlDados.attributes.getNamedItem("VALOR").value
									if iSeq = 1 then
										sDescCol = xmlDados.attributes.getNamedItem("DESCRICAO").value
										sColunas = sColunas & "<td class='td_tabelas_titulo centro' style='white-space: nowrap' >"&sDescCol&"</td>"
									end if
								end if
		
								'Código
								if xmlDados.nodeName  = "CODIGO" then
									sCodigo = xmlDados.attributes.getNamedItem("VALOR").value
								end if
							
								'Quantidade
								if xmlDados.nodeName  = "QUANTIDADE" then
									sQuantidade = xmlDados.attributes.getNamedItem("VALOR").value
									if iSeq = 1 then
										sDescCol = xmlDados.attributes.getNamedItem("DESCRICAO").value
										sColunas = sColunas & "<td class='td_tabelas_titulo centro' style='white-space: nowrap; width: 45px'>"&sDescCol&"</td>"
									end if
								end if
								
								'Situação
								if xmlDados.nodeName  = "SITUACAO" then
									sSituacao = xmlDados.attributes.getNamedItem("VALOR").value
									if iSeq = 1 then
										sDescCol = xmlDados.attributes.getNamedItem("DESCRICAO").value
										sColunas = sColunas & "<td class='td_tabelas_titulo centro' style='width: 120px'>"&sDescCol&"</td>"
									end if
								end if
								
								'Prioridade
								if xmlDados.nodeName  = "PRIORIDADE" then
									sPrioridade = xmlDados.attributes.getNamedItem("VALOR").value
									if iSeq = 1 then
										sDescCol = xmlDados.attributes.getNamedItem("DESCRICAO").value
										sColunas = sColunas & "<td class='td_tabelas_titulo centro' style='white-space: nowrap; width: 70px'>"&sDescCol&"</td>"
									end if
								end if								
							Next
	
							if iSeq = 1 then
								'Grid de sugestões									
								Response.write "<table class='max_width' style='border-spacing: 1px'>"
								Response.write "<tr style='height: 20px'>"
								Response.Write sColunas
								Response.write "</tr>"
							end if

							'Define estilo da linha
							if (sSeq mod 2) > 0 then 'Linha ímpar
								td_class = "td_tabelas_valor2"
							else 'Linha par
								td_class = "td_tabelas_valor1"
							end if

							 Response.write "<tr style='height: 20px'>"
							 Response.write "<td class='centro "&td_class&"'>"&sSeq&"</td>"
							 Response.write "<td style='width: 80px' class='centro "&td_class&"'>"&sData&"</td>"
							 link_det_sgt = "index.asp?content=aquisicoes&acao=detalhe&codsgt="&sCodigo&"&codusu="&codigo_usu&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma
							 Response.write "<td class='esquerda "&td_class&"'><a class='link_serv' href='"&link_det_sgt&"'>&nbsp;"&sTitulo&"</a></td>"
							 Response.write "<td class='centro "&td_class&"'>&nbsp;"&sEdicao&"&nbsp;</td>"
							 Response.write "<td class='esquerda "&td_class&"'>&nbsp;"&sAutor&"&nbsp;</td>"
							 Response.write "<td class='centro "&td_class&"'>&nbsp;"&sEditora&"&nbsp;</td>"
							 Response.write "<td class='centro "&td_class&"'>"&sQuantidade&"</td>"
							 
							 if sSituacao <> "" then
								 Response.write "<td class='centro "&td_class&"'>"&sSituacao&"</td>"
							 end if

							 Response.write "<td class='centro "&td_class&"'>"&sPrioridade&"</td>"															 
							 
							 Response.write "</tr>"
		 				end if
					Next
					
					
					Response.write "</table><br />"
				else
				   Response.write "<table class='max_width' style='border-spacing: 1px'>"
				   Response.write "<tr>"
				   Response.write "<td class='centro td_tabelas_valor2' colspan=8>"
				   msg = getTermo(global_idioma, 1338, "Não existem sugestões para %s no momento.", 0)
				   msg = Format(msg, "<b>"&Formata_Nome(Session("nome_usuario"),"inteiro")&"</b>")
				   Response.write msg
				   Response.write "</td></tr>"
				   Response.write "</table>"
				end if				
			end if
		end if
	end if
%>
</td>
</tr>
</table>