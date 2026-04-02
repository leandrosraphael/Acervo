<link href="../inc/contraste.css" rel="stylesheet" type="text/css" />

<%
if veio_de_popup = 1 then
	doc_root = "../"
else
	veio_de_popup = 0
	doc_root = ""
end if

'****************************************
' EXIBE OS EXEMPLARES
'****************************************

if left(xml_exemplar,5) = "<?xml" then
	infoReservas = ""
	estiloAno = ""
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml xml_exemplar
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "FICHA" then
		For Each xmlPNode In xmlRoot.childNodes
			'****************************************
			' MONTA GRID DE COLEÇÕES (UNICAMP)
			'****************************************
			if Tipo = "0" then
				if xmlPNode.nodeName = "COLECOES" then
					sequencial = 0
					colunas_grid = ""
					reg_colecao = ""
					
					'****************************************
					' CADA COLEÇÃO
					'****************************************
					For Each xmlColecao In xmlPNode.childNodes
						sequencial = sequencial + 1
						
						if ((sequencial mod 2) > 0) then
							'---------- IMPAR
							fontcolor = "black"
							td_class = "td_tabelas_valor1"
							link_class = "link_serv"
						else
							'------------ PAR
							fontcolor= "#000000"
							td_class = "td_tabelas_valor2"
							link_class = "link_serv"
						end if
						
						colecao_atual = "<tr style='height: 25px'>"
			
						if xmlColecao.nodeName = "COLECAO" then
							For Each xmlCampos In xmlColecao.childNodes
								'****************************************
								' BIBLIOTECA
								'****************************************
								if xmlCampos.nodeName = "BIBLIOTECA" then
									valor = xmlCampos.attributes.getNamedItem("VALOR").value
									if (sequencial = 1) then
										descricao = xmlCampos.attributes.getNamedItem("DESCRICAO").value
										colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>&nbsp;"&descricao&"&nbsp;</td>"
									end if
									cod_bib = xmlCampos.attributes.getNamedItem("CODIGO_BIBLIOTECA").value
									descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(cod_bib) & "," & veio_de_popup & ",0);'>"&valor&"</a>"
									
									colecao_atual = colecao_atual & "<td class='centro "&td_class&"'>"&descricaoBib&"</td>"
								end if
											
								'****************************************
								' COLEÇÃO IMPRESSA
								'****************************************
								if xmlCampos.nodeName = "COL_IMPRESSA" then
									if (sequencial = 1) then
										descricao = xmlCampos.attributes.getNamedItem("DESCRICAO").value
										colunas_grid = colunas_grid & "<td class='esquerda td_tabelas_titulo'>&nbsp;"&descricao&"&nbsp;</td>"
									end if

									valor = xmlCampos.attributes.getNamedItem("VALOR").value
									colecao_atual = colecao_atual & "<td class='esquerda "&td_class&"'>&nbsp;"&valor&"</td>"
								end if
								
								'****************************************
								' COMPLEMENTO
								'****************************************
								if xmlCampos.nodeName = "COMPLEMENTO" then
									if (sequencial = 1) then
										descricao = xmlCampos.attributes.getNamedItem("DESCRICAO").value
										colunas_grid = colunas_grid & "<td class='esquerda td_tabelas_titulo'>&nbsp;"&descricao&"&nbsp;</td>"
									end if
									
									valor = xmlCampos.attributes.getNamedItem("VALOR").value
									colecao_atual = colecao_atual & "<td class='esquerda "&td_class&"'>&nbsp;"&valor&"</td>"
								end if
							Next
						end if
						
						colecao_atual = colecao_atual & "</tr>"
						reg_colecao = reg_colecao & colecao_atual
					Next
					
					NumColecao = sequencial
					
					'****************************************
					' MONTA O GRID DE COLEÇÃO
					'****************************************
					Response.Write "<table class='table-exemplares'>"
					Response.Write "<tr style='height: 20px'>"
					Response.Write colunas_grid
					Response.Write "</tr>"
					Response.Write reg_colecao
					Response.Write "</table><br />"
				end if
			
		
			'****************************************
			' MONTA COMBO DE ANOS PARA OS PERIODICOS
			'****************************************
			
				if xmlPNode.nodeName = "ANOS" then
					comboanos = ""
					For Each xmlAnos In xmlPNode.childNodes
						'****************************************
						' ANO SELECIONADO
						'****************************************
						if xmlAnos.nodeName = "ANO_ATUAL" then
							if Request.QueryString("ano") = "" then
								ano_atual = xmlAnos.attributes.getNamedItem("Valor").value
							else
								if Request.QueryString("ano") = "_BRANCO_" then
									ano_atual = ""
								else
									ano_atual = Request.QueryString("ano")
								end if
							end if	
						end if
						'****************************************
						' COMBO COM OS ANOS
						'****************************************
						if xmlAnos.nodeName = "ANO"	then
							if (trim(xmlAnos.attributes.getNamedItem("Valor").value) = "") then
								valor_ano = "_BRANCO_"
							else
								valor_ano = xmlAnos.attributes.getNamedItem("Valor").value
							end if

							if ano_atual = xmlAnos.attributes.getNamedItem("Valor").value then
								comboanos = comboanos & "<option value='"&valor_ano&"' selected>"&xmlAnos.attributes.getNamedItem("Valor").value&"</option>"
							else
								comboanos = comboanos & "<option value='"&valor_ano&"'>"&xmlAnos.attributes.getNamedItem("Valor").value&"</option>"
							end if
						end if
					Next
					
					if comboanos <> "" then
						if ano_atual = "" then
							comboanos = "<option value='TODOS' selected>" & getTermo(global_idioma, 318, "Todos", 0) & "</option>" & comboanos
						else
							comboanos = "<option value='TODOS'>" & getTermo(global_idioma, 318, "Todos", 0) & "</option>" & comboanos
						end if
					end if
				end if
			end if
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
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo' style=' width:20px'>&nbsp;"&descricao&"&nbsp;</td>"
					end if
					'****************************************
					' CAPA DO FASCICULO
					'****************************************
					if xmlColunas.nodeName = "CAPA_FASCICULO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>&nbsp;"&descricao&"&nbsp;</td>"
					end if
					'****************************************
					' CODIGO
					'****************************************
					if xmlColunas.nodeName = "CODEX" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' TOMBO
					'****************************************
					if xmlColunas.nodeName = "TOMBO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' EDIÇÃO
					'****************************************
					if xmlColunas.nodeName = "EDICAO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' ANO
					'****************************************
					if xmlColunas.nodeName = "ANO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo' #ESTILO_ANO#>"&descricao&"</td>"
					end if
					'****************************************
					' VOLUME
					'****************************************
					if xmlColunas.nodeName = "VOLUME" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' NUMERO
					'****************************************
					if xmlColunas.nodeName = "NUMERO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' PARTE
					'****************************************
					if xmlColunas.nodeName = "PARTE" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' PERIODO DE CIRCULACAO
					'****************************************
					if xmlColunas.nodeName = "PERC_CIRC" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo' title='"&getTermo(global_idioma, 1163, "Período de circulação", 0)&"'>"&descricao&"</td>"
					end if
					'****************************************
					' SUPORTE
					'****************************************
					if xmlColunas.nodeName = "SUPORTE" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' ASSINATURA TOPOGRÁFICA
					'****************************************
					if xmlColunas.nodeName = "ASSINATURA_TOPOGRAFICA" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if					
					'****************************************
					' DATA DE PUBLICACAO
					'****************************************
					if xmlColunas.nodeName = "DATA_PUBLICACAO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo' title='"&getTermo(global_idioma, 10, "Data de publicação", 0)&"'>"&descricao&"</td>"
					end if
					'****************************************
					' NUMERO DE CHAMADA
					'****************************************
					if xmlColunas.nodeName = "NUM_CHAMADA" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo' style='white-space: nowrap'>"&descricao&"</td>"
					end if
					'****************************************
					' CAMPO OPCIONAL
					'****************************************
					if xmlColunas.nodeName = "CMP_OPCIONAL" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='td_tabelas_titulo'>&nbsp;"&descricao&"&nbsp;</td>"
					end if
					'****************************************
					' TABELA OPCIONAL
					'****************************************
					if xmlColunas.nodeName = "TAB_OPCIONAL" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' NOTAS
					'****************************************
					if xmlColunas.nodeName = "NOTAS" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>&nbsp;"&descricao&"&nbsp;</td>"
					end if
					'****************************************
					' SUBLOCALIZACAO
					'****************************************
					if xmlColunas.nodeName = "SUBLOCALIZACAO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' BIBLIOTECA
					'****************************************
					if xmlColunas.nodeName = "BIBLIOTECA" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'****************************************
					' SITUAÇÃO
					'****************************************
					if xmlColunas.nodeName = "SITUACAO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'>"&descricao&"</td>"
					end if
					'******************************************
					' NOTAS DE EXEMPLAR - EXIBIÇÃO CONFIGURAVEL
					'******************************************
					if xmlCampos.nodeName = "NOTA_EXEMPLAR" then
						colunas_grid = colunas_grid  & "<td class='centro td_tabelas_titulo'>"&xmlColunas.attributes.getNamedItem("DESCRICAO").value&"</td>"
					end if
					'****************************************
					' ARTIGOS
					'****************************************
					if xmlColunas.nodeName = "ARTIGOS" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo' style='width: 40px'>"&descricao&"</td>"
					end if
					'****************************************
					' RESERVA
					'****************************************
					if xmlColunas.nodeName = "RESERVA" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo' style='width: 20px'>"&descricao&"</td>"
					end if
					'****************************************
					' MIDIA NO FASCÍCULO
					'****************************************
					if xmlColunas.nodeName = "MIDIA_FASCICULO" then
						descricao =  xmlColunas.attributes.getNamedItem("Descricao").value
						colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo'><span style='font-size: 20px; color: #0070c0; line-height: 5px;'>"&descricao&"</td>"
					end if
				Next

				'****************************************
				' QR CODE
				'****************************************
				if (global_cfg_hab_qrcode_exemplar = 1) then
					colunas_grid = colunas_grid & "<td class='centro td_tabelas_titulo' style='width: 60px'>"&getTermo(global_idioma, 9871, "QR Code", 0)&"</td>"
				end if
			end if
			
			'****************************************
			' MONTA GRID COM TOTAIS DE EXEMPLARES (UNISANTANNA)
			'****************************************
			if xmlPNode.nodeName = "TOTAIS_EXEMPLARES_BIBLIOTECA" then
					ExibeTotalPorBiblioteca = true
					sequencial = 1
					tabela_resultado = ""

					For Each xmlTotalExemplar In xmlPNode.childNodes
					
						if (sequencial mod 2) > 0 then 
							'---------- IMPAR
							fontcolor = "black" 	
							td_class = "td_tabelas_valor1"
							link_class = "link_serv"
						else 
							'------------ PAR
							fontcolor= "#000000" 
							td_class = "td_tabelas_valor2"
							link_class = "link_serv"								
						end if	
					
						if xmlTotalExemplar.nodeName = "TOTAL_EXEMPLARES_BIBLIOTECA" then
							iColspan = 0
							For Each xmlCampos In xmlTotalExemplar.childNodes
								valor = ""
								'Capturando os dados da biblioteca
								if xmlCampos.nodeName = "BIBLIOTECA" then
									
									'Biblioteca do exemplar
									valor = xmlCampos.attributes.getNamedItem("BIBLIOTECA").value
									if (valor <> "") then
										tabela_resultado = tabela_resultado & "<tr><td class='centro td_tabelas_titulo' style=' width: 200px' rowspan='4'>&nbsp;" & valor & "&nbsp;</td></tr>"
									end if
									
									'Exemplares disponíveis
									valor = xmlCampos.attributes.getNamedItem("NUM_EXS_DISP").value
									if (valor <> "") then
										tabela_resultado = tabela_resultado & "<tr>"
										descricao = "Exemplares disponíveis: "
										tabela_resultado = tabela_resultado & "<td class='esquerda td_tabelas_titulo' style=' width: 200px'>&nbsp;"&descricao&"&nbsp;</td>"
										tabela_resultado = tabela_resultado & "<td class='centro "&td_class&"'style='width: 100px'>&nbsp;"&valor&"&nbsp;</td>"
										tabela_resultado = tabela_resultado & "</tr>"
									end if
									
									'Exemplares para consulta
									valor = xmlCampos.attributes.getNamedItem("NUM_EXS_CONS").value
									if (valor <> "") then
										tabela_resultado = tabela_resultado & "<tr>"
										descricao = "Exemplares para consulta: "
										tabela_resultado = tabela_resultado & "<td class='esquerda td_tabelas_titulo' style=' width: 200px'>&nbsp;"&descricao&"&nbsp;</td>"
										tabela_resultado = tabela_resultado & "<td class='centro "&td_class&"'>&nbsp;"&valor&"&nbsp;</td>"
										tabela_resultado = tabela_resultado & "</tr>"
									end if
									
									'Reservas
									valor = xmlCampos.attributes.getNamedItem("NUM_EXS_RESE").value
									if (valor <> "") then
										tabela_resultado = tabela_resultado & "<tr>"
										descricao = "Reservas: "
										tabela_resultado = tabela_resultado & "<td class='esquerda td_tabelas_titulo' style='width: 200px'>&nbsp;"&descricao&"&nbsp;</td>"
										tabela_resultado = tabela_resultado & "<td class='centro "&td_class&"'>&nbsp;"&valor&"&nbsp;</td>"
										tabela_resultado = tabela_resultado & "</tr>"
									end if

									sequencial = sequencial + 1									
								end if
								
								if (valor <> "") then
									iColsPan = iColsPan + 1
								end if
							next
						end if
					next
					
					if (tabela_resultado <> "") then
						'****************************************
						' MONTA O GRID DE TOTAIS DE EXEMPLARES
						'****************************************
						Response.write "<table class='table-exemplares'>"
						Response.write tabela_resultado
						Response.write "</table><br />"
					end if					
			end if

			'****************************************
			' INFORMAÇÕES DE RESERVAS
			'****************************************
			if xmlPNode.nodeName = "RESERVAS" then
				For Each xmlReservas In xmlPNode.childNodes	
					reservaBiblioteca = xmlReservas.attributes.getNamedItem("Biblioteca").value
					reservaQtde = xmlReservas.attributes.getNamedItem("Qtde").value

					if CInt(reservaQtde) <> 0 then
						infoReservas = infoReservas & "<li style='list-style: disc;'>" & reservaBiblioteca & ": " & "<b>" & reservaQtde & "</b> "

						if CInt(reservaQtde) = 1 then
							infoReservas = infoReservas & getTermo(global_idioma, 348, "Reserva", 2)
						else
							infoReservas = infoReservas & getTermo(global_idioma, 347, "Reservas", 2)
						end if

						infoReservas = infoReservas & ";</li>"
					end if
				Next
				if infoReservas <> "" then
					infoReservas = "<ul style='margin-top: 5px; margin-left: 35px; list-style: disc;'>" & infoReservas & "</ul>"
				end if
			end if

			'****************************************
			' GRID DOS EXEMPLARES
			'****************************************
			if xmlPNode.nodeName = "EXEMPLARES" then
				sequencial = 1
				reg_exemplares = ""
				NumExemplares = xmlPNode.attributes.getNamedItem("NUM_EXS").value
				NumReservas = xmlPNode.attributes.getNamedItem("QTDE_RESERVA").value
				
				'**************************************************************************************
				' Adequação Belas Artes - Exibe informações adicionais, totalizando as situações
				'**************************************************************************************
				NumExDisp = xmlPNode.attributes.getNamedItem("NUM_EXS_DISP").value
				NumExEmp = xmlPNode.attributes.getNamedItem("NUM_EXS_EMP").value
				NumExIndisp = xmlPNode.attributes.getNamedItem("NUM_EXS_INDISP").value
				
				For Each xmlExemplares In xmlPNode.childNodes
					if (sequencial mod 2) > 0 then 
						'---------- IMPAR
						fontcolor = "black" 	
						td_class = "td_tabelas_valor1"
						link_class = "link_serv"
					else 
						'------------ PAR
						fontcolor= "#000000" 
						td_class = "td_tabelas_valor2"
						link_class = "link_serv"								
					end if	
					'****************************************
					' CADA EXEMPLAR
					'****************************************
					if xmlExemplares.nodeName = "EXEMPLAR" then
						exemplar_atual = "<tr style=' height: 25px'>"
						For Each xmlCampos In xmlExemplares.childNodes
							'****************************************
							' SEQUENCIAL
							'****************************************
							if xmlCampos.nodeName = "SEQ" then
								valor = sequencial
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' CAPA DO FASCICULO
							'****************************************
							if xmlCampos.nodeName = "CAPA_FASCICULO" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								if (valor <> "") then							
									exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'><img src=data:image/png;base64,"&valor&" /></td>"
								else
									exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'></td>"
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
							' EDIÇÃO
							'****************************************
							if xmlCampos.nodeName = "EDICAO" then
								valor_ed = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor_ed&"</td>"
							end if
							'****************************************
							' ANO
							'****************************************
							if xmlCampos.nodeName = "ANO" then
								valor_ano = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor_ano&"</td>"
								if len(valor_ano) > 10 then
									estiloAno = "style='min-width: 150px;'"
								end if
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
							' ASSINATURA TOPOGRÁFICA
							'****************************************
							if xmlCampos.nodeName = "ASSINATURA_TOPOGRAFICA" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if							
							'****************************************
							' DATA DE PUBLICACAO
							'****************************************
							if xmlCampos.nodeName = "DATA_PUBLICACAO" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' NUMERO DE CHAMADA
							'****************************************
							if xmlCampos.nodeName = "NUM_CHAMADA" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								if CStr(valor) = "" then
									exemplar_atual = exemplar_atual & "<td class='esquerda "&td_class&"'>&nbsp;-&nbsp;</td>"
								else
									exemplar_atual = exemplar_atual & "<td class='esquerda "&td_class&"'>"&valor&"</td>"
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
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' NOTAS
							'****************************************
							if xmlCampos.nodeName = "NOTAS" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"' style='text-align: left; padding-left: 5px;'>&nbsp;"&valor&"&nbsp;</td>"
							end if
							'****************************************
							' SUBLOCALIZACAO
							'****************************************
							if xmlCampos.nodeName = "SUBLOCALIZACAO" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' BIBLIOTECA
							'****************************************
							if xmlCampos.nodeName = "BIBLIOTECA" then
								cod_bib = xmlCampos.attributes.getNamedItem("Codigo").value
								cod_bib_atual = xmlCampos.attributes.getNamedItem("CODIGO_BIB_ATUAL").value
								if (cod_bib = cod_bib_atual) then
									biblioteca = xmlCampos.attributes.getNamedItem("NOME_BIBLIOTECA").value
									if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB").value = 1) then
										descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(cod_bib) & "," & veio_de_popup & ",0);'>"&biblioteca&"</a>"
									else
										descricaoBib = biblioteca
									end if
								else
									bibliotecaAtual = xmlCampos.attributes.getNamedItem("NOME_BIB_ATUAL").value
									biblioteca = xmlCampos.attributes.getNamedItem("NOME_BIBLIOTECA").value 
									if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB").value = 1) then
										descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(cod_bib) & "," & veio_de_popup & ",0);'>"&biblioteca&"</a>"
									else
										descricaoBib = biblioteca
									end if
									descricaoBib = descricaoBib & "<br />"
									if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB_ATUAL").value = 1) then
										descricaoBib = descricaoBib & "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(cod_bib_atual) & "," & veio_de_popup & ",0);'>(" & getTermo(global_idioma, 4558, "Emp.", 0)& " " & bibliotecaAtual&")</a>"
									else
										descricaoBib = "(" & getTermo(global_idioma, 4558, "Emp.", 0)& " " & bibliotecaAtual & ")"
									end if
								end if
								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&descricaoBib&"</td>"
							end if
							'****************************************
							' SITUAÇÃO
							'****************************************
							if xmlCampos.nodeName = "SITUACAO" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								sCodEx = xmlCampos.attributes.getNamedItem("Exemplar").value
								solic_emp = xmlCampos.attributes.getNamedItem("SolicEmp").value

								if (repositorio_institucional = 0) and (solic_emp = "1") and (Session("usuario_externo") = false) then
									valor = valor & "<br/>"
									valor = valor & "<a href='javascript:SolicitarEmprestimo(" & sCodEx & ", " & iIndexSrv & ", " & veio_de_popup & ");' class='link_classic'>"
									valor = valor & getTermo(global_idioma, 7451, "Solicitar empréstimo", 0)
									valor = valor & "</a>"
								end if

								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'******************************************
							' NOTAS DE EXEMPLAR - EXIBIÇÃO CONFIGURAVEL
							'******************************************
							if xmlCampos.nodeName = "NOTA_EXEMPLAR" then
								codigoExemplar = xmlCampos.attributes.getNamedItem("EXEMPLAR").value
								possuiNota = xmlCampos.attributes.getNamedItem("POSSUI_NOTA").value
								
								if (possuiNota > 0) then
									valor = "<a href='javascript:visualizarNotaExemplar(" & codigoExemplar & ", " & iIndexSrv & ", " & veio_de_popup & ");' class='link_classic'>"&getTermo(global_idioma, 2660, "ver", 2)&"</a>"
								else
									valor = "&nbsp;-&nbsp;"
								end if

								exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&valor&"</td>"
							end if
							'****************************************
							' ARTIGOS
							'****************************************
							if xmlCampos.nodeName = "ARTIGOS" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								if CStr(valor) = "0" then
									exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>&nbsp;-&nbsp;</td>"
								else
									'Pega o código do exemplar, para montar o link para os artigos ligados a ele
									sCodEx = xmlCampos.attributes.getNamedItem("Exemplar").value
									if veio_de_popup = 1 then
										sArt_Link = "<a href=""javascript:parent.abrePopup2('asp/artigos.asp?popup=1&codex="&sCodEx&"&Servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 1339, "Artigos", 0)&"',965, 450, false, true);"" title='"&getTermo(global_idioma, 1339, "Artigos", 0)&"' class='link_classic'>"&getTermo(global_idioma, 2660, "ver", 2)&"</a>"
									else
										sArt_Link = "<a href=""javascript:abrePopup('asp/artigos.asp?codex="&sCodEx&"&Servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 1339, "Artigos", 0)&"', 965, 450, false, true);"" title='"&getTermo(global_idioma, 1339, "Artigos", 0)&"' class='link_classic'>"&getTermo(global_idioma, 2660, "ver", 2)&"</a>"
									end if
									exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&sArt_Link&"</td>"
								end if
							end if
							'****************************************
							' RESERVA
							'****************************************
							if xmlCampos.nodeName = "RESERVA" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								if CStr(valor) = "1" then
									if Session("Logado") = "sim" then
										exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'><span class='transparent-icon span_imagem icon_16 icon-small-flag-b-h' title='"&getTermo(global_idioma, 2661, "Reservar exemplar", 0)&"' style='cursor:pointer;' onclick=""ReservaEx(1,'"&Session("codigo_usuario")&"',"&CStr(Tipo)&",'"&codigo_obra&"','"&Trim(valor_ano)&"','"&Trim(valor_volume)&"','"&Trim(valor_num)&"','"&Trim(cod_bib)&"','"&doc_root&"',"&iIndexSrv&");""></span></td>"
									else
										exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'><span class='span_imagem icon_16 icon-small-flag-x' title='"&getTermo(global_idioma, 2661, "Reservar exemplar", 0)&"' style='cursor:pointer;' onclick=""ReservaEx(0,'0',"&CStr(Tipo)&",'"&codigo_obra&"','"&Trim(valor_ano)&"','"&Trim(valor_volume)&"','"&Trim(valor_num)&"','"&Trim(cod_bib)&"','"&doc_root&"',"&iIndexSrv&");""></span></td>"
									end if
								else
									exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>x</td>"
								end if
							end if
							'*******************
							' MIDIA NO FASCÍCULO
							'*******************
							if xmlCampos.nodeName = "MIDIA_FASCICULO" then
								valor = xmlCampos.attributes.getNamedItem("Valor").value
								if CStr(valor) = "0" then
									exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>&nbsp;-&nbsp;</td>"
								else
									sCodEx = xmlCampos.attributes.getNamedItem("Exemplar").value
									 
									if veio_de_popup = 1 then
										sArt_Link = "<a href=""javascript:parent.abrePopup2('asp/midia_exemplar.asp?tipo="&Tipo&"&codex="&sCodEx&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 7517, "Conteúdo digital", 0)&"',500,490,true,true);"" title='"&getTermo(global_idioma, 7517, "Conteúdo digital", 0)&"' class='link_classic'>"&getTermo(global_idioma, 2660, "ver", 2)&"</a>"
									else
										sArt_Link = "<a href=""javascript:abrePopup('asp/midia_exemplar.asp?tipo="&Tipo&"&codex="&sCodEx&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 7517, "Conteúdo digital", 0)&"',500,490,true,true);"" title='"&getTermo(global_idioma, 7517, "Conteúdo digital", 0)&"' class='link_classic'>"&getTermo(global_idioma, 2660, "ver", 2)&"</a>"
									end if
									exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&sArt_Link&"</td>"
								end if
							end if

							'****************************************
							' CODIGO DO EXEMPLAR
							'****************************************
							if xmlCampos.nodeName = "CODIGO_EXEMPLAR" then
								sCodigoExemplar = xmlCampos.attributes.getNamedItem("Valor").value
							end if
						Next

						'*******************
						' QR Code
						'*******************
						if (global_cfg_hab_qrcode_exemplar = 1) then	

							if veio_de_popup = 1 then
								qr_code = "<a href=""javascript:parent.abrePopup2('asp/informacao_exemplar_qrcode.asp?tipo="&Tipo&"&codex="&sCodigoExemplar&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 9872, "Informações do exemplar", 0)&"',730,260,true,true);"" title='"&getTermo(global_idioma, 9872, "Informações do exemplar", 0)&"' class='link_classic'><span class='span_imagem div_imagem_right_3 icon_16 icon-small-qrcode' data-icon='qrcode'></span></a>"
							else
								qr_code = "<a href=""javascript:abrePopup('asp/informacao_exemplar_qrcode.asp?tipo="&Tipo&"&codex="&sCodigoExemplar&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 9872, "Informações do exemplar", 0)&"',730,260,true,true);"" title='"&getTermo(global_idioma, 9872, "Informações do exemplar", 0)&"' class='link_classic'><span class='span_imagem div_imagem_right_3 icon_16 icon-small-qrcode' data-icon='qrcode'></span></a>"
							end if

       						exemplar_atual = exemplar_atual & "<td class='centro "&td_class&"'>"&qr_code&"</td>"
						end if

						exemplar_atual = exemplar_atual & "</tr>"
						'****************************************
						' FILTRAR POR ANO SE FOR PERIODICO
						'****************************************

						if CStr(Tipo) = "0" then
							if Request.QueryString("filtra_artigo") = "" then 
								if valor_ano = ano_atual or ano_atual = "#TODOS#" or ano_atual = "TODOS" then
									reg_exemplares = reg_exemplares & exemplar_atual
									sequencial = sequencial + 1
								end if
							else
								reg_exemplares = reg_exemplares & exemplar_atual
								sequencial = sequencial + 1
							end if
						else
							reg_exemplares = reg_exemplares & exemplar_atual
							sequencial = sequencial + 1
						end if
					end if				
				Next
			end if
		Next
	end if
	Set xmlDoc = nothing
	Set xmlRoot = nothing
else
	if global_numero_serie = 3156 then 'SE FOR ITA NÂO EXIBE MENSAGEM "Nenhum exemplar disponível!" 
		response.write "<br />"
	else
		response.write getTermo(global_idioma, 2662, "Nenhum exemplar disponível", 0) 
	end if
end if

'****************************************
' NÃO EXISTEM EXEMPLARES
'****************************************
if (NumExemplares <= 0) AND (NumColecao <= 0) AND (not ExibeTotalPorBiblioteca) then
	if global_numero_serie = 3156 then 'SE FOR ITA NÂO EXIBE MENSAGEM "Nenhum exemplar disponível!" 
		response.write "<br />"
	else
		response.write getTermo(global_idioma, 2662, "Nenhum exemplar disponível", 0)&"<br />" 
	end if
else
	if (NumColecao <= 0) OR ((NumExemplares > 0) AND ((global_numero_serie = 6387) or (global_numero_serie = 4134) or (global_numero_serie = 6235) or (global_numero_serie = 5516) or (global_numero_serie = 7869))) then        
		if veio_de_popup = 1 then
			espaco = "&nbsp;&nbsp;"
		else
			espaco = "&nbsp;&nbsp;&nbsp;&nbsp;"
		end if
	
		'**************************************************************************************
		' Adequação Belas Artes - Exibe informações adicionais, totalizando as situações
		'**************************************************************************************
		sInfoComp = ""
		if CStr(NumExDisp) <> "" then
			sInfoComp = sInfoComp & espaco & getTermo(global_idioma, 1808, "Nº de exemplares disponíveis", 0) & ": <b>" & NumExDisp & "</b><br />"
		end if
		if CStr(NumExEmp) <> "" then
			sInfoComp = sInfoComp & espaco & getTermo(global_idioma, 2663, "Nº de exemplares emprestados", 0) & ": <b>" & NumExEmp & "</b><br />"
		end if
		if CStr(NumExIndisp) <> "" then
			sInfoComp = sInfoComp & espaco & getTermo(global_idioma, 2517, "Nº de exemplares indisponíveis (retido/não circula)", 0) & ": <b>" & NumExIndisp & "</b><br />"
		end if
	
		'****************************************
		' EXIBE QTDE DOS EXEMPLARES E RESERVAS
		' E MONTA O COMBO DOS ANOS
		'****************************************
		if CStr(Tipo) = "0" then
			if NumReservas = 0 then
				sReserva = espaco & getTermo(global_idioma, 2238, "Não existem reservas para este periódico", 0)
			elseif NumReservas = 1 then
				msg_res = getTermo(global_idioma, 2141, "Existe %s para este periódico", 0)
				msg_res = Format(msg_res, "<span style='color:red'><b>1</b>" & " " & getTermo(global_idioma, 348, "Reserva", 2) & "</span>")
				sReserva = espaco & msg_res
			else
				msg_res = getTermo(global_idioma, 2219, "Existem %s para este periódico", 0)
				msg_res = Format(msg_res, "<span style='color:red'><b>"&NumReservas&"</b>" & " " & getTermo(global_idioma, 347, "Reservas", 2) & "</span>")
				sReserva = espaco & msg_res
			end if
			if infoReservas <> "" then
				sReserva = sReserva & ":" & infoReservas
			end if
				
			Response.write "<form name='frm_detalhe' method='post' action='#' target='mainFrame'>"
			if Request.QueryString("filtra_artigo") = "" then
				Response.write "<div class='esquerda'>" & espaco & getTermo(global_idioma, 2666, "Total de fascículos que a biblioteca possui", 0) & ": <b>"&NumExemplares&"</b><br />" & sInfoComp & sReserva & "</div><p class='centro'>"
				estilo_combo_ano = "style='width: 100px;'"
				if(global_numero_serie = 5592) then
					estilo_combo_ano = "style='width: 180px;'"
				end if
				if veio_de_popup = 1 then
					Response.write "<br />"&getTermo(global_idioma, 1849, "Exemplares de", 0)&" <select class='select_logica styled_combo' " & estilo_combo_ano & " id='ano' name='ano' onchange=envia_combo(parent.parent.hiddenFrame.modo_busca,"&codigo_obra&","&request.QueryString("qtde")&","&request.QueryString("pagina")&","&request.QueryString("posicao_vetor")&","&iIndexSrv&")>"
				else
					Response.write "<br />"&getTermo(global_idioma, 1849, "Exemplares de", 0)&" <select class='select_logica styled_combo' " & estilo_combo_ano & " id='ano' name='ano' onchange=envia_combo(parent.hiddenFrame.modo_busca,"&codigo_obra&","&request.QueryString("qtde")&","&request.QueryString("pagina")&","&request.QueryString("posicao_vetor")&","&iIndexSrv&")>"
				end if
				Response.Write comboanos
			else
				if (bFiltroPorBiblioteca) then
					if (bTermoBibSingular) then
						sTextoLink = getTermo(global_idioma, 8402, "Clique aqui para ver os fascículos da biblioteca selecionada.", 0)
					else
						sTextoLink = getTermo(global_idioma, 8403, "Clique aqui para ver os fascículos das bibliotecas selecionadas.", 0)
					end if
					sTextoTitle = getTermo(global_idioma, 8404, "Ver os fascículos deste periódico.", 0)
				else
					sTextoLink = getTermo(global_idioma, 1850, "Clique aqui para ver todos.", 0)
					sTextoTitle = getTermo(global_idioma, 1851, "Ver todos os fascículos deste periódico", 0)
				end if
				if veio_de_popup = 1 then
					ver_todos_exemp = "<a class='link_classic' title='"&sTextoTitle&"' href=""javascript:LinkDetalhes(parent.parent.hiddenFrame.modo_busca,1,1,"&codigo_obra&",1,'periodico',0);"">"&sTextoLink&"</a>"
				else
					ver_todos_exemp = "<a class='link_classic' title='"&sTextoTitle&"' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1,"&codigo_obra&",1,'periodico',0);"">"&sTextoLink&"</a>"
				end if
				msg_bibex = getTermo(global_idioma, 1914, "A biblioteca possui %s fascículos deste periódico.", 0)
				msg_bibex = Format(msg_bibex, "<b>"&NumExemplares&"</b>")
				Response.write "<table><tr><td class='td_filtro_exemplar'>"&msg_bibex&" "&ver_todos_exemp&"</tr></td></table><br />"
			end if
			Response.write "</select></p>"
			Response.write "<input type='hidden' name='codigo_obra' value="&codigo_obra&">"
			Response.write "</form><br />"
		'****************************************
		' EXIBE QTDE DOS EXEMPLARES E RESERVAS
		'****************************************
		elseif CStr(Tipo) = "1" then
			if NumReservas = 0 then
				sReserva = espaco & getTermo(global_idioma, 2518, "Não existem reservas para esta obra", 0)
			elseif NumReservas = 1 then
				msg_res = getTermo(global_idioma, 2664, "Existe %s para esta obra", 0)
				msg_res = Format(msg_res, "<span style='color:red'><b>1</b>" & " " & getTermo(global_idioma, 348, "Reserva", 2) & "</span>")
				sReserva = espaco & msg_res
			else
				msg_res = getTermo(global_idioma, 2665, "Existem %s para esta obra", 0)
				msg_res = Format(msg_res, "<span style='color:red'><b>"&NumReservas&"</b>" & " " & getTermo(global_idioma, 347, "Reservas", 2) & "</span>")
				sReserva = espaco & msg_res
			end if
			if infoReservas <> "" then
				sReserva = sReserva & ":" & infoReservas
			end if
			if (not ExibeTotalPorBiblioteca) then
				Response.write "<div class='esquerda'>" & espaco & getTermoHtml(global_idioma, 2170, "Nº de exemplares", 0) & ": <b>" & NumExemplares
				Response.write "</b><br />" & sInfoComp & sReserva & "</div><br />"
			end if
		end if
		
		'****************************************
		' MONTA O GRID DE EXEMPLAR
		'****************************************
		Response.write "<table class='table-exemplares'>"
		Response.write "<tr style='height: 20px'>"
		Response.Write replace(colunas_grid, "#ESTILO_ANO#", estiloAno)
		Response.write "</tr>"
		Response.write reg_exemplares
		Response.write "</table><br />"
	end if
end if
%>