<%
	response.write "<table class='remover_bordas_padding max_width' style='background-color: #CCCCCC'>"
	response.write "<tr style='height: 26px'>"
	response.write "<td style='text-align: left;'>"
	response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a class='link_serv' href=""javascript:LinkAquisicoes(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;"&getTermo(global_idioma, 1386, "voltar", 2)&"</a>"
	response.write "</td>"
	response.write "</tr>"
	Response.write "</table>"
	
	response.write "<table class='max_width'>"
	response.write "<tr>"
	response.write "<td class='td_padrao' style='text-align: center; vertical-align: top'>"
	response.write "<br/>"
	
	'Busca os dados da sugestão
	Set ROService = ROServer.CreateService("Web_Consulta")	
	sXML = ROService.DetalheSugestao(request.querystring("codsgt"),global_idioma)
	Set ROService = nothing
	TrataErros(1)

	if (left(sXML,5) = "<?xml") then
		'Processa o XML com os dados da sugestão
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml sXML
		Set xmlRoot = xmlDoc.documentElement

		'Verifica se usuário possui sugestões
		if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then	
			For Each xmlDados In xmlRoot.childNodes	
				'Data
				if xmlDados.nodeName  = "DATA" then
					sData_T = xmlDados.attributes.getNamedItem("DESCRICAO").value					
					sData   = xmlDados.attributes.getNamedItem("VALOR").value					
				end if
			
				'Titulo
				if xmlDados.nodeName  = "TITULO" then
					sTitulo_T =  xmlDados.attributes.getNamedItem("DESCRICAO").value
					sTitulo   =  xmlDados.attributes.getNamedItem("VALOR").value
				end if
				
				'Prioridade
				if xmlDados.nodeName  = "PRIORIDADE" then
					sPrioridade_T =  xmlDados.attributes.getNamedItem("DESCRICAO").value
					sPrioridade   =  xmlDados.attributes.getNamedItem("VALOR").value
				end if				

				'Edição
				if xmlDados.nodeName  = "EDICAO" then
					sEdicao_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sEdicao   = xmlDados.attributes.getNamedItem("VALOR").value					
				end if

				'Autor
				if xmlDados.nodeName  = "AUTOR" then
					sAutor_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sAutor   = xmlDados.attributes.getNamedItem("VALOR").value					
				end if
				
				'Editora
				if xmlDados.nodeName  = "EDITORA" then
					sEditora_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sEditora   = xmlDados.attributes.getNamedItem("VALOR").value					
				end if

				'Quantidade
				if xmlDados.nodeName  = "QUANTIDADE" then
					sQuantidade_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sQuantidade = xmlDados.attributes.getNamedItem("VALOR").value
				end if
				
				'Volume
				if xmlDados.nodeName  = "VOLUME" then
					sVolume_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sVolume   = xmlDados.attributes.getNamedItem("VALOR").value
				end if

				'Ano
				if xmlDados.nodeName  = "ANO" then
					sAno_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sAno   = xmlDados.attributes.getNamedItem("VALOR").value
				end if
				
				'Biblioteca
				if xmlDados.nodeName  = "BIBLIOTECA" then
					sBiblioteca_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sBiblioteca   = xmlDados.attributes.getNamedItem("VALOR").value
                    sCodigoBib = xmlDados.attributes.getNamedItem("CODIGO_BIBLIOTECA").value
				end if
				
				'ISBN
				if xmlDados.nodeName  = "ISBN" then
					sISBN_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sISBN   = xmlDados.attributes.getNamedItem("VALOR").value
				end if
				
				'Nacional
				if xmlDados.nodeName  = "NACIONAL" then
					sNacional_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sNacional   = xmlDados.attributes.getNamedItem("VALOR").value
				end if
				
				'Observações
				if xmlDados.nodeName  = "OBSERVACOES" then
					sObs_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sObs   = xmlDados.attributes.getNamedItem("VALOR").value
				end if

				'Curso
				if xmlDados.nodeName  = "CURSO" then
					sCurso_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sCurso   = xmlDados.attributes.getNamedItem("VALOR").value
				end if
				
				'Série
				if xmlDados.nodeName  = "SERIE" then
					sSerie_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sSerie   = xmlDados.attributes.getNamedItem("VALOR").value
				end if
				
				'Disciplina
				if xmlDados.nodeName  = "DISCIPLINA" then
					sDisciplina_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sDisciplina   = xmlDados.attributes.getNamedItem("VALOR").value
				end if
				
				'Bibliografia
				if xmlDados.nodeName  = "BIBLIOGRAFIA" then
					sBibliografia_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sBibliografia   = xmlDados.attributes.getNamedItem("VALOR").value
				end if
				
				'Material
				if xmlDados.nodeName  = "MATERIAL" then
					sMaterial_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sMaterial   = xmlDados.attributes.getNamedItem("VALOR").value
				end if

				'Valor
				if xmlDados.nodeName  = "VALOR" then
					sValor_T = xmlDados.attributes.getNamedItem("DESCRICAO").value
					sValor   = xmlDados.attributes.getNamedItem("VALOR").value
				end if
			Next
			
			'----------------------------------------------------------
			'Carrega os valores
	 		response.write "<table style='width: 90%; border-spacing: 2px; padding: 2px; margin: auto;'>"	
  			response.write "<tr><td colspan='2' class='td_detalhe_descricao' style='text-align: center; height: 23px'>&nbsp;<b>"&getTermo(global_idioma, 1442, "Detalhes da sugestão", 0)&"</b>&nbsp;</td></tr>"
			
			'Data da sugestão
			response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sData_T&"</td>"
			response.write "<td class='td_detalhe_valor'>&nbsp;"&sData&"</td></tr>"
			'Titulo
			response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sTitulo_T&"</td>"
			response.write "<td class='td_detalhe_valor'><b>&nbsp;"&sTitulo&"</b></td></tr>"
			'Prioridade
			response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sPrioridade_T&"</td>"
			response.write "<td class='td_detalhe_valor'>&nbsp;"&sPrioridade&"</td></tr>"
			'Autor
			if len(trim(sAutor)) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sAutor_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sAutor&"</td></tr>"
			end if
			'Material
			if len(trim(sMaterial)) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sMaterial_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sMaterial&"</td></tr>"
			end if
			'Editoras
			if len(trim(sEditora)) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sEditora_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sEditora&"</td></tr>"
			end if
			'ISBN
			if len(trim(sISBN)) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sISBN_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sISBN&"</td></tr>"
			end if
			'Ano
			if len(trim(sAno)) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sAno_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sAno&"</td></tr>"
			end if
			'Edicao
			if len(trim(sEdicao)) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sEdicao_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sEdicao&"</td></tr>"
			end if
			'Volume
			if len(trim(sVolume)) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sVolume_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sVolume&"</td></tr>"	
			end if
			'Biblioteca
			if len(trim(sBiblioteca)) > 0 then
                descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(sCodigoBib) & ",0,0);'>"&sBiblioteca&"</a>"
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sBiblioteca_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&descricaoBib&"</td></tr>"	
			end if
			'Quantidade
			if len(trim(sQuantidade)) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sQuantidade_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sQuantidade&"</td></tr>"
			end if
			'Valor Estimado
			if sValor > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sValor_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sValor&"</td></tr>"
			end if		
			'Nacional
			response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sNacional_T&"</td>"
			response.write "<td class='td_detalhe_valor'>&nbsp;"&sNacional&"</td></tr>"
			'Curso
			if len(sCurso) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sCurso_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sCurso&"</td></tr>"		
			end if
			'Série
			if len(sSerie) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sSerie_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sSerie&"</td></tr>"		
			end if
			'Disciplina
			if len(sDisciplina) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sDisciplina_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sDisciplina&"</td></tr>"		
			end if
			'Bibliografia
			if len(sBibliografia) > 0 then
				response.write "<tr><td style='text-align: left;' class='td_detalhe_descricao'>&nbsp;"&sBibliografia_T&"</td>"
				response.write "<td class='td_detalhe_valor'>&nbsp;"&sBibliografia&"</td></tr>"		
			end if	
			
			response.write "</table><br/>"
		
			'Observações
			if len(trim(sObs)) > 0 then 
				response.write "<table style='width: 90%; border-spacing: 2px; padding: 2px; margin: auto;'>"
				response.write "<tr><td class='td_detalhe_descricao' style='text-align: center; height: 23px'><b>"&sObs_T&"</b>&nbsp;</td></tr>"
				response.write "<tr><td class='td_detalhe_valor' style='text-align: left;'>&nbsp;"&trim(replace(sObs,chr(10),"<br/>"))&"</td></tr>"
				response.write "</table><br/>"
			end if
		end if
	end if
	response.write "</td></tr></table>"
%>