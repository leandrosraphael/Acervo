<%
'------------------------------------------------------------
'------------------------ F O N T E S -----------------------
'------------------------------------------------------------
NumFontes = 0

if left(xml_exemplar,5) = "<?xml" then
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml xml_exemplar
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "FICHA" then
		For Each xmlPNode In xmlRoot.childNodes
			'****************************************
			' MONTA CABEÇALHO DO GRID DE EXEMPLARES
			'****************************************
			if xmlPNode.nodeName = "COLUNAS" then
				colunas_grid = ""
				For Each xmlColunas In xmlPNode.childNodes
					'****************************************
					' SEQUENCIAL
					'****************************************
					if xmlColunas.nodeName = "SEQ" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao' style='width: 20px'>&nbsp;"&descricao&"&nbsp;</td>"
					end if
					'****************************************
					' TITULO
					'****************************************
					if xmlColunas.nodeName = "TITULO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' CODIGO
					'****************************************
					if xmlColunas.nodeName = "CODEX" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' TOMBO
					'****************************************
					if xmlColunas.nodeName = "TOMBO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' ANO
					'****************************************
					if xmlColunas.nodeName = "ANO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' VOLUME
					'****************************************
					if xmlColunas.nodeName = "VOLUME" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' NUMERO
					'****************************************
					if xmlColunas.nodeName = "NUMERO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' PARTE
					'****************************************
					if xmlColunas.nodeName = "PARTE" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' PERIODO DE CIRCULACAO
					'****************************************
					if xmlColunas.nodeName = "PERC_CIRC" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao' title='"&getTermo(global_idiomas, 1163, "Período de circulação", 0)&"'>"&descricao&"</td>"
					end if
					'****************************************
					' SUPORTE
					'****************************************
					if xmlColunas.nodeName = "SUPORTE" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' NUMERO DE CHAMADA
					'****************************************
					if xmlColunas.nodeName = "NUM_CHAMADA" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao' style='white-space: nowrap'>"&descricao&"</td>"
					end if
					'****************************************
					' CAMPO OPCIONAL
					'****************************************
					if xmlColunas.nodeName = "CMP_OPCIONAL" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>&nbsp;"&descricao&"&nbsp;</td>"
					end if
					'****************************************
					' TABELA OPCIONAL
					'****************************************
					if xmlColunas.nodeName = "TAB_OPCIONAL" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' BIBLIOTECA
					'****************************************
					if xmlColunas.nodeName = "BIBLIOTECA" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' SITUAÇÃO
					'****************************************
					if xmlColunas.nodeName = "SITUACAO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_legislacao'>"&descricao&"</td>"
					end if
					'****************************************
					' RESERVA
					'****************************************
					if (global_esconde_menu = 0) then 
						if xmlColunas.nodeName = "RESERVA" then
							descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
							colunas_grid = colunas_grid & "<td class='centro td_legislacao' style=' width: 20px'>"&descricao&"</td>"
						end if
					end if
				Next
			end if
			'****************************************
			' GRID DAS FONTES
			'****************************************
			if xmlPNode.nodeName = "EXEMPLARES" then
				sequencial = 1
				reg_exemplares = ""
				NumFontes = xmlPNode.attributes.getNamedItem("NUM_EXS").value
				NumReservas = xmlPNode.attributes.getNamedItem("QTDE_RESERVA").value
				For Each xmlExemplares In xmlPNode.childNodes
					if (sequencial mod 2) > 0 then 
						'---------- IMPAR
						td_class = "td_tabelas_valor2"
						link_class = "link_serv"
					else 
						'------------ PAR
						td_class = "td_tabelas_valor1"
						link_class = "link_serv"								
					end if	
					'****************************************
					' CADA FONTE
					'****************************************
					if xmlExemplares.nodeName = "EXEMPLAR" then
						exemplar_atual = "<tr style='height: 19px'>"
						For Each xmlCampos In xmlExemplares.childNodes
							'****************************************
							' SEQUENCIAL
							'****************************************
							if xmlCampos.nodeName = "SEQ" then
								valor = sequencial
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' TITULO
							'****************************************
							if xmlCampos.nodeName = "TITULO" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								codigo_obra = xmlCampos.attributes.getNamedItem("Codigo").value
								if (global_numero_serie = 4794) and (global_somente_legislacao = 1) then
									link_tit = "<a class="&link_class&" href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1,"&codigo_obra&",1,'leg',0);"">"&valor&"</a>"
									exemplar_atual = exemplar_atual & "<td class='esquerda "&td_class&"'>&nbsp;"&link_tit&"&nbsp;</td>"
								else
									exemplar_atual = exemplar_atual & "<td class='esquerda "&td_class&"'>"&valor&"</td>"
								end if
								
							end if
							'****************************************
							' CODIGO
							'****************************************
								if xmlCampos.nodeName = "CODEX" then
									valor = xmlCampos.attributes.getNamedItem("Valor").value
									exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
								end if
							'****************************************
							' TOMBO
							'****************************************
							if xmlCampos.nodeName = "TOMBO" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' ANO
							'****************************************
							if xmlCampos.nodeName = "ANO" then
								valor_ano = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor_ano&"</td>"
							end if
							'****************************************
							' VOLUME
							'****************************************
							if xmlCampos.nodeName = "VOLUME" then
								valor_volume = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor_volume&"</td>"
							end if
							'****************************************
							' NUMERO
							'****************************************
							if xmlCampos.nodeName = "NUMERO" then
								valor_num = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor_num&"</td>"
							end if
							'****************************************
							' PARTE
							'****************************************
							if xmlCampos.nodeName = "PARTE" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' PERIODO DE CIRCULACAO
							'****************************************
							if xmlCampos.nodeName = "PERC_CIRC" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' SUPORTE
							'****************************************
							if xmlCampos.nodeName = "SUPORTE" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' NUMERO DE CHAMADA
							'****************************************
							if xmlCampos.nodeName = "NUM_CHAMADA" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								if CStr(valor) = "" then
									exemplar_atual = exemplar_atual & "<td class='esquerda "&td_class&"' style='white-space: nowrap'>&nbsp;-&nbsp;</td>"
								else
									exemplar_atual = exemplar_atual & "<td class='esquerda "&td_class&"' style='white-space: nowrap'>"&valor&"</td>"
								end if
							end if
							'****************************************
							' CAMPO OPCIONAL
							'****************************************
							if xmlCampos.nodeName = "CMP_OPCIONAL" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>&nbsp;"&valor&"&nbsp;</td>"
							end if
							'****************************************
							' TABELA OPCIONAL
							'****************************************
							if xmlCampos.nodeName = "TAB_OPCIONAL" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='esquerda "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' BIBLIOTECA
							'****************************************
							if xmlCampos.nodeName = "BIBLIOTECA" then
								cod_bib = xmlCampos.attributes.getNamedItem("Codigo").value
								valor_bib = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor_bib&"</td>"
							end if
							'****************************************
							' SITUAÇÃO
							'****************************************
							if xmlCampos.nodeName = "SITUACAO" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' RESERVA
							'****************************************
							if (global_esconde_menu = 0) then 
								if xmlCampos.nodeName = "RESERVA" then
									valor = xmlCampos.attributes.getNamedItem("Valor").value
									if CStr(valor) = "1" then
										if Session("Logado") = "sim" then
											exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"' ><span class='transparent-icon span_imagem icon_16 icon-small-flag-b-h'  title='"&getTermo(global_idioma, 2661, "Reservar exemplar", 0)&"' style='cursor:pointer;' onclick=""ReservaEx(1,'"&Session("codigo_usuario")&"',"&CStr(Tipo)&",'"&codigo_obra&"','"&Trim(valor_ano)&"','"&Trim(valor_volume)&"','"&Trim(valor_num)&"','"&Trim(cod_bib)&"','',"&iIndexSrv&");""></span></td>"
										else
											exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'><span class='span_imagem icon_16 icon-small-flag-x' title='"&getTermo(global_idioma, 2661, "Reservar exemplar", 0)&"' style='cursor:pointer;' onclick=""ReservaEx(0,'0',"&CStr(Tipo)&",'"&codigo_obra&"','"&Trim(valor_ano)&"','"&Trim(valor_volume)&"','"&Trim(valor_num)&"','"&Trim(cod_bib)&"','',"&iIndexSrv&");""></span></td>"
										end if
									else
										exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>x</td>"
									end if
								end if
							end if
						Next
						exemplar_atual = exemplar_atual & "</tr>"
						reg_exemplares = reg_exemplares & exemplar_atual
						sequencial = sequencial + 1
					end if				
				Next
			end if
		Next
	end if
	Set xmlDoc = nothing
	Set xmlRoot = nothing
end if
		
if NumFontes > 0 then
	'****************************************
	' MONTA O GRID DE EXEMPLAR
	'****************************************
	Response.write "<span style='size: 2'>"&getTermo(global_idioma, 2701, "Nº de fontes", 0)&":&nbsp;<b>" & NumFontes & "</b></span><br /><br />"
	Response.write "<table style=' width: 100%; padding: 1; border-spacing: 1px; background-color: #CCCCCC'>"
	Response.write "<tr style='height: 20px'>"
	Response.Write colunas_grid
	Response.write "</tr>"
	Response.write reg_exemplares
	Response.write "</table><br/>"
end if

%>