<%
'--------------------------------------------------------------------------------
' 				MONTA RESULTADO EM FORMATO DE FICHA
'--------------------------------------------------------------------------------
Response.Write "<table class='tabela_grid_sels' style='border-spacing: 0; padding: 1px'><tr>"
	Response.Write sXMLFichas
if (left(sXMLFichas,5) = "<?xml") then
	sequencial = 1
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml sXMLFichas
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "Pagina" then
		For Each xmlFicha In xmlRoot.childNodes
			if xmlFicha.nodeName  = "Ficha" then
				Tipo = xmlFicha.attributes.getNamedItem("Tipo").value
				Registro = xmlFicha.attributes.getNamedItem("Registro").value
				
				Response.Write "<tr>"
				Response.Write "<td class='td_center_top td_grid_ficha_marg' style='width: 40px; white-space:nowrap; border-spacing: 1px; padding: 0; padding-top: 5px;'>"&Registro&"</td>"
				'//-------------------------------------------------------------//
				'---------------------CELULA COM A FICHA------------------------//
				'//-------------------------------------------------------------//				
				Response.Write "<td class='td_center_middle td_grid_ficha_marg'>"
				ficha = ""
				codigo_atual = 0
				tem_chamada = false
								
				For Each xmlCampos In xmlFicha.childNodes
					'*************************************
					'FICHA DE PERIÓDICOS
					'*************************************
					if cStr(Tipo) = "0" then
						'*************************************
						'CODIGO
						'*************************************
						if xmlCampos.nodeName = "CODIGO" then
							codigo_atual = xmlCampos.attributes.getNamedItem("Valor").value
						end if
						'*************************************
						'MATERIAL
						'*************************************
						if xmlCampos.nodeName = "MIDIA" then
							material = xmlCampos.attributes.getNamedItem("Valor").value
							material_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='direita td_imp_esq'>"
							ficha = ficha & material_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&material&"</td></tr>"
						end if
						'*************************************
						'NUMERO DE CHAMADA
						'*************************************
						if xmlCampos.nodeName = "CHAMADA" then
							chamada = xmlCampos.attributes.getNamedItem("Valor").value
							chamada_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & chamada_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&chamada&"</td></tr>"
						end if
						'*************************************
						'ANDAR (UNICENP)
						'*************************************
						if xmlCampos.nodeName = "ANDAR" then
							andar = xmlCampos.attributes.getNamedItem("Valor").value
							andar_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & andar_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&andar&"</td></tr>"
						end if
                        '*************************************
						'LOCALIZAÇÃO (BN)
						'*************************************
						if xmlCampos.nodeName = "LOCALIZACAO" then
							loc = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							loc_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & loc_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&loc&"</td></tr>"
						end if
						'*************************************
						'TITULO
						'*************************************
						if xmlCampos.nodeName = "TITULO" then
							titulo = xmlCampos.attributes.getNamedItem("Titulo").value
							titulo_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							complemento = " " & xmlCampos.attributes.getNamedItem("Complemento").value
							titulo_f = Replace(Replace(titulo,"<","&lt;"),">","&gt;")
							total = GetSessao(global_cookie,"nrows"&iIndexSrv)
							pos_dec = Registro - 1
						
 							ficha = ficha & "<tr>"
 							ficha = ficha & "<td class='td_imp_esq direita'>"
 							ficha = ficha & titulo_desc
 							ficha = ficha & "</td><td class='td_imp_dir esquerda'><b>"
 							ficha = ficha & titulo_f&"</b>"&complemento
 							ficha = ficha & "</td></tr>"	
						end if
						'*************************************
						'OUTROS TITULOS (Adequação UNIUBE)
						'*************************************
						if xmlCampos.nodeName = "OUTROS_TITULOS" then
							For Each xmlSubCampos In xmlCampos.childNodes
								desc_outros_titulos = trim(xmlSubCampos.attributes.getNamedItem("Descricao").value)
								valor_outros_titulos = trim(xmlSubCampos.attributes.getNamedItem("Valor").value)
								
								ficha = ficha & "<tr>"
								ficha = ficha & "<td class='td_imp_esq direita'>"
								ficha = ficha & desc_outros_titulos
								ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&valor_outros_titulos&"</td></tr>"
							Next
						end if
                        '*************************************
						'ANO
						'Adequeação IMS: não mostrar
						'*************************************
						if xmlCampos.nodeName = "ANO" and (global_numero_serie <> 4516) then
							ano = xmlCampos.attributes.getNamedItem("Valor").value
							ano_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & ano_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&ano&"</td></tr>"
						end if
						'************************************************************
						' IMPRENTA
						'Adequeação: apenas IMS
						'************************************************************
						if xmlCampos.nodeName = "IMPRENTA" and (global_numero_serie = 4516) then
							imprenta = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_imprenta = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & desc_imprenta
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&imprenta&"</td></tr>"
						end if
						'************************************************************
						' Nº ÁLBUM
						'Adequeação: apenas IMS
						'************************************************************
						if xmlCampos.nodeName = "COMPLEMENTO1" and (global_numero_serie = 4516) then
							complemento = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_complemento = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & desc_complemento
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&complemento&"</td></tr>"
						end if
						'*************************************
						'PERIODICIDADE
						'*************************************
						if xmlCampos.nodeName = "PERIODICIDADE" then
							periodicidade = xmlCampos.attributes.getNamedItem("Valor").value
							periodicidade_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & periodicidade_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&periodicidade&"</td></tr>"
						end if
						'*************************************
						'ASSUNTO
						'*************************************
						if xmlCampos.nodeName = "ASSUNTOS" then
							seq = 1
							assunto = ""
							exibe_todos_ass = trim(xmlCampos.attributes.getNamedItem("ExibeTodos").value)
							desc_assunto = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'*************************************
								'SE EXIBE TODOS OS ASSUNTOS NA FICHA
								'*************************************
								if (exibe_todos_ass = "1") then
									tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
									desc_princ_assunto = replace(replace(trim(xmlSubCampos.attributes.getNamedItem("Desc_Princ").value)," ","_"),"&#39;","_#39;")
									valor_assunto = trim(xmlSubCampos.attributes.getNamedItem("Valor").value)
									seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)

									temp = seq_assunto&" "&replace(valor_assunto,chr(10)," <br/> ")

									if assunto <> "" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & "&nbsp;" & temp & "&nbsp;"
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='autLink td_left_middle'>" & temp & "&nbsp;</td></tr></table>"
										end if
									else
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & temp & "&nbsp;"
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='autLink td_left_middle'>" & temp & "&nbsp;</td></tr></table>"
										end if
									end if
									seq = seq + 1
								'*************************************
								'SE EXIBE SOMENTE PRIMEIRO ASSUNTO
								'*************************************
								else
									desc_assunto = xmlSubCampos.attributes.getNamedItem("Valor").value
									assunto_desc = xmlSubCampos.attributes.getNamedItem("Descricao").value
									assunto_princ = xmlSubCampos.attributes.getNamedItem("Desc_Princ").value
									assunto_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							
									assunto_formatado = replace(replace(replace(assunto_princ," ","_"),"<",""),">","")
																	
									ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&desc_assunto&"</td></tr>"
								end if
							Next
							if (exibe_todos_ass = "1") then
								ficha = ficha & "<td class='td_imp_dir esquerda'>" & assunto & "</td>"
								ficha = ficha & "</tr>"
							end if
						end if
						'*************************************
						'SITES
						'*************************************
						if xmlCampos.nodeName = "SITE" then
							site_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							site_c = ""
							For Each xmlItem In xmlCampos.childNodes
								site = TrocaTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
								site_url = RemoveTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
								if site_c <> "" then
									site_c = site_c + "<br/>"
								end if
								site_c = site_c + "<p>" & site & "</p>"
							Next	

							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & site_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>" & site_c & "</td></tr>"
						end if
					'*************************************
					'FICHA DE OBRAS
					'*************************************
					elseif cStr(Tipo) = "1" then
						'*************************************
						'CODIGO
						'*************************************
						if xmlCampos.nodeName = "CODIGO" then
							codigo_atual = xmlCampos.attributes.getNamedItem("Valor").value
						end if
						'*************************************
						'MATERIAL
						'*************************************
						if xmlCampos.nodeName = "MIDIA" then
							material = xmlCampos.attributes.getNamedItem("Valor").value
							material_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & material_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&material&"</td></tr>"
						end if
						'*************************************
						'NUMERO DE CHAMADA
						'*************************************
						if xmlCampos.nodeName = "CHAMADA" then
							chamada = xmlCampos.attributes.getNamedItem("Valor").value
							chamada_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & chamada_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&chamada&"</td></tr>"
						end if
						'*************************************
						'ANDAR (UNICENP)
						'*************************************
						if xmlCampos.nodeName = "ANDAR" then
							andar = xmlCampos.attributes.getNamedItem("Valor").value
							andar_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & andar_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&andar&"</td></tr>"
						end if
                        '*************************************
						'LOCALIZAÇÃO (BN)
						'*************************************
						if xmlCampos.nodeName = "LOCALIZACAO" then
							loc = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							loc_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & loc_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&loc&"</td></tr>"
						end if
						'*************************************
						'AUTOR (ENTRADA PRINCIPAL)
						'*************************************
						if xmlCampos.nodeName = "ENT_PRINC" then
							desc_autor = xmlCampos.attributes.getNamedItem("Valor").value
							autor_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							autor_princ = xmlCampos.attributes.getNamedItem("Desc_Princ").value
							autor_tipo = xmlCampos.attributes.getNamedItem("Tipo").value
							autor_codigo = xmlCampos.attributes.getNamedItem("Codigo").value							
							
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & autor_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&desc_autor&"</td></tr>"
						end if
						'*************************************
						'AUTORIA (ENTRADA PRINCIPAL e ENTRADA(S) SECUNDÁRIA(S))
						'Adequação para o cliente IMS
						'*************************************
						if xmlCampos.nodeName = "AUTORIA" then
							autor = ""
							autor_desc = xmlCampos.attributes.getNamedItem("Desc_Autoria").value
							sequencialOld = sequencial
							
							For Each xmlSubCampos In xmlCampos.childNodes
								sequencial = sequencial + 1
								desc_autor = xmlSubCampos.attributes.getNamedItem("Valor").value
								autor_princ = xmlSubCampos.attributes.getNamedItem("Desc_Princ").value
								autor_tipo = xmlSubCampos.attributes.getNamedItem("Tipo").value
								autor_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							
																								
								autor = autor & desc_autor & " <br/> "
							Next
							sequencial = sequencialOld
								
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & autor_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&autor&"</td></tr>"
						end if
						'*************************************
						'TITULO
						'*************************************
						if xmlCampos.nodeName = "TITULO" then
							titulo = xmlCampos.attributes.getNamedItem("Titulo").value
							titulo_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							complemento = " " & xmlCampos.attributes.getNamedItem("Complemento").value
							titulo_f = Replace(Replace(titulo,"<","&lt;"),">","&gt;")
							total = GetSessao(global_cookie,"nrows"&iIndexSrv)
							pos_dec = Registro - 1							
													
 							ficha = ficha & "<tr>"
 							ficha = ficha & "<td class='td_imp_esq direita'>"
 							ficha = ficha & titulo_desc
 							ficha = ficha & "</td><td class='td_imp_dir esquerda'><b>"
 							ficha = ficha & titulo_f&"</b>"&complemento
 							ficha = ficha & "</td></tr>"	
						end if
                        '*************************************
						'ANO
						'Adequeação IMS: não mostrar
						'*************************************
						if xmlCampos.nodeName = "ANO" and (global_numero_serie <> 4516) then
							ano = xmlCampos.attributes.getNamedItem("Valor").value
							ano_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & ano_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&ano&"</td></tr>"
						end if
						'************************************************************
						' IMPRENTA
						'Adequeação: apenas IMS
						'************************************************************
						if xmlCampos.nodeName = "IMPRENTA" and (global_numero_serie = 4516) then
							imprenta = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_imprenta = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & desc_imprenta
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&imprenta&"</td></tr>"
						end if
						'************************************************************
						' Nº ÁLBUM
						'Adequeação: apenas IMS
						'************************************************************
						if xmlCampos.nodeName = "COMPLEMENTO1" and (global_numero_serie = 4516) then
							complemento = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_complemento = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & desc_complemento
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&complemento&"</td></tr>"
						end if
						'*************************************
						'DATA
						'*************************************
						if xmlCampos.nodeName = "DATA" then
							data = xmlCampos.attributes.getNamedItem("Valor").value
							data_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & data_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&data&"</td></tr>"
						end if
						'*************************************
						'ASSUNTO
						'*************************************
						if xmlCampos.nodeName = "ASSUNTOS" then
							seq = 1
							assunto = ""
							exibe_todos_ass = trim(xmlCampos.attributes.getNamedItem("ExibeTodos").value)
							desc_assunto = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'*************************************
								'SE EXIBE TODOS OS ASSUNTOS NA FICHA
								'*************************************
								if (exibe_todos_ass = "1") then
									tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
									desc_princ_assunto = replace(replace(trim(xmlSubCampos.attributes.getNamedItem("Desc_Princ").value)," ","_"),"&#39;","_#39;")
									valor_assunto = trim(xmlSubCampos.attributes.getNamedItem("Valor").value)
									seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)

									temp = seq_assunto&" "&replace(valor_assunto,chr(10)," <br/> ")

									if assunto <> "" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & "&nbsp;" & temp & "&nbsp;"
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0' ><tr><td id='assLinkTab"&seq&"' class='autLink td_left_middle'>" & temp & "&nbsp;</td></tr></table>"
										end if
									else
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & temp & "&nbsp;"
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='autLink td_left_middle'>" & temp & "&nbsp;</td></tr></table>"
										end if
									end if
									seq = seq + 1
								'*************************************
								'SE EXIBE SOMENTE PRIMEIRO ASSUNTO
								'*************************************
								else
									desc_assunto = xmlSubCampos.attributes.getNamedItem("Valor").value
									assunto_desc = xmlSubCampos.attributes.getNamedItem("Descricao").value
									assunto_princ = xmlSubCampos.attributes.getNamedItem("Desc_Princ").value
									assunto_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							
									assunto_formatado = replace(replace(replace(assunto_princ," ","_"),"<",""),">","")
																	
									ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&desc_assunto&"</td></tr>"
								end if
							Next
							if (exibe_todos_ass = "1") then
								ficha = ficha & "<td class='td_imp_dir esquerda'>" & assunto & "</td>"
								ficha = ficha & "</tr>"
							end if
						end if
						'*************************************
						'SITES
						'*************************************
						if xmlCampos.nodeName = "SITE" then
							site_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							site_c = ""
							For Each xmlItem In xmlCampos.childNodes
								site = TrocaTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
								site_url = RemoveTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
								if site_c <> "" then
									site_c = site_c + "<br/>"
								end if
								site_c = site_c + "<p>" & site & "</p>"
							Next	

							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & site_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>" & site_c & "</td></tr>"
						end if				
					'*************************************
					'FICHA DE LEGISLAÇÃO
					'*************************************						
					elseif cStr(Tipo) = "2" then
						'*************************************
						'CODIGO
						'*************************************
						if xmlCampos.nodeName = "CODIGO" then
							codigo_atual = xmlCampos.attributes.getNamedItem("Valor").value
						end if
						'*************************************
						'MATERIAL
						'*************************************
						if xmlCampos.nodeName = "MIDIA" then
							material = xmlCampos.attributes.getNamedItem("Valor").value
							material_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & material_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&material&"</td></tr>"
						end if
						'*************************************
						'NORMA E NÚMERO
						'*************************************
						if xmlCampos.nodeName = "NORMA" then
							norma = xmlCampos.attributes.getNamedItem("Valor").value
							norma_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							norma_f = Replace(Replace(norma,"<","&lt;"),">","&gt;")
							total = GetSessao(global_cookie,"nrows"&iIndexSrv)
							pos_dec = Registro - 1

 							ficha = ficha & "<tr>"
 							ficha = ficha & "<td class='td_imp_esq direita'>"
 							ficha = ficha & norma_desc
 							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"
 							ficha = ficha & norma_f
 							ficha = ficha & "</td></tr>"	
						end if
						'*************************************
						'APELIDO
						'*************************************
						if xmlCampos.nodeName = "APELIDO" then
							apelido = xmlCampos.attributes.getNamedItem("Valor").value
							apelido_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & apelido_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&apelido&"</td></tr>"
						end if
						'*************************************
						'DATA ASSINATURA
						'*************************************
						if xmlCampos.nodeName = "DATA_ASSINATURA" then
							data_assinatura = xmlCampos.attributes.getNamedItem("Valor").value
							data_assinatura_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & data_assinatura_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&data_assinatura&"</td></tr>"
						end if
						'*************************************
						'EMENTA
						'*************************************
						if xmlCampos.nodeName = "EMENTA" then
							ementa = xmlCampos.attributes.getNamedItem("Valor").value
							ementa_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & ementa_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&ementa&"</td></tr>"
						end if
						'*************************************
						'DATA DE PUPLICAÇÃO
						'*************************************
						if xmlCampos.nodeName = "DATA_PUBLICACAO" then
							data_publicacao = xmlCampos.attributes.getNamedItem("Valor").value
							data_publicacao_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & data_publicacao_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&data_publicacao&"</td></tr>"
						end if						
						'*************************************
						'ORGAO DE ORIGEM
						'*************************************
						if xmlCampos.nodeName = "ORGAO_ORIGEM" then
							desc_orgao = xmlCampos.attributes.getNamedItem("Valor").value
							orgao_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							orgao_princ = xmlCampos.attributes.getNamedItem("Desc_Princ").value
							orgao_codigo = xmlCampos.attributes.getNamedItem("Codigo").value							
							orgao_formatado = replace(replace(replace(orgao_princ," ","_"),"<",""),">","")
							
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & orgao_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&desc_orgao&"</td></tr>"
						end if
						'*************************************
						'ESFERA
						'*************************************
						if xmlCampos.nodeName = "ESFERA" then
							esfera = xmlCampos.attributes.getNamedItem("Valor").value
							esfera_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq esquerda'>"
							ficha = ficha & esfera_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&esfera&"</td></tr>"
						end if
						'*************************************
						'SITUAÇÃO DA LEGISLAÇÃO
						'*************************************
						if xmlCampos.nodeName = "SITUACAO_LEGISLACAO" then
							leg_situacao = xmlCampos.attributes.getNamedItem("Valor").value
							leg_situacao_desc = xmlCampos.attributes.getNamedItem("Descricao").value

							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & leg_situacao_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&leg_situacao&"</td></tr>"
                        end if
						'*************************************
						'ASSUNTO
						'*************************************
						if xmlCampos.nodeName = "ASSUNTOS" then
							seq = 1
							assunto = ""
							exibe_todos_ass = trim(xmlCampos.attributes.getNamedItem("ExibeTodos").value)
							desc_assunto = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'*************************************
								'SE EXIBE TODOS OS ASSUNTOS NA FICHA
								'*************************************
								if (exibe_todos_ass = "1") then
									tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
									desc_princ_assunto = replace(replace(trim(xmlSubCampos.attributes.getNamedItem("Desc_Princ").value)," ","_"),"&#39;","_#39;")
									valor_assunto = trim(xmlSubCampos.attributes.getNamedItem("Valor").value)
									seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)

									temp = seq_assunto&" "&replace(valor_assunto,chr(10)," <br/> ")

									if assunto <> "" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & "&nbsp;" & temp & "&nbsp;"
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='autLink td_left_middle'>" & temp & "&nbsp;</td></tr></table>"
										end if
									else
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & temp & "&nbsp;"
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='autLink td_left_middle'>" & temp & "&nbsp;</td></tr></table>"
										end if
									end if
									seq = seq + 1
								'*************************************
								'SE EXIBE SOMENTE PRIMEIRO ASSUNTO
								'*************************************
								else
									desc_assunto = xmlSubCampos.attributes.getNamedItem("Valor").value
									assunto_desc = xmlSubCampos.attributes.getNamedItem("Descricao").value
									assunto_princ = xmlSubCampos.attributes.getNamedItem("Desc_Princ").value
									assunto_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							
									assunto_formatado = replace(replace(replace(assunto_princ," ","_"),"<",""),">","")
																	
									ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&desc_assunto&"</td></tr>"
								end if
							Next
							if (exibe_todos_ass = "1") then
								ficha = ficha & "<td class='td_imp_dir esquerda'>" & assunto & "</td>"
								ficha = ficha & "</tr>"
							end if
						end if
						'*************************************
						'SITES
						'*************************************
						if xmlCampos.nodeName = "SITE" then
							site_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							site_c = ""
							For Each xmlItem In xmlCampos.childNodes
								site = TrocaTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
								site_url = RemoveTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
								if site_c <> "" then
									site_c = site_c + "<br/>"
								end if
								site_c = site_c + "<p>" & site & "</p>"
							Next	

							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & site_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>" & site_c & "</td></tr>"
						end if
					'*************************************
					'FICHA DE ANALITICA
					'*************************************
					elseif cStr(Tipo) = "3" then
						'*************************************
						'CODIGO
						'*************************************
						if xmlCampos.nodeName = "CODIGO" then
							codigo_atual = xmlCampos.attributes.getNamedItem("Valor").value
						end if
						'*************************************
						'MATERIAL
						'*************************************
						if xmlCampos.nodeName = "MIDIA" then
							material = xmlCampos.attributes.getNamedItem("Valor").value
							material_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & material_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&material&"</td></tr>"
						end if
						'*************************************
						'NUMERO DE CHAMADA
						'*************************************
						if xmlCampos.nodeName = "CHAMADA" then
							chamada = xmlCampos.attributes.getNamedItem("Valor").value
							chamada_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & chamada_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&chamada&"</td></tr>"
						end if
						'*************************************
						'AUTOR (ENTRADA PRINCIPAL)
						'*************************************
						if xmlCampos.nodeName = "ENT_PRINC" then
							desc_autor = xmlCampos.attributes.getNamedItem("Valor").value
							autor_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							autor_princ = xmlCampos.attributes.getNamedItem("Desc_Princ").value
							autor_tipo = xmlCampos.attributes.getNamedItem("Tipo").value
							autor_codigo = xmlCampos.attributes.getNamedItem("Codigo").value					
									
							ficha = ficha & "<tr>"
							ficha = ficha & "<td colspan=2 class='td_imp_esq direita'>"
							ficha = ficha & autor_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&desc_autor&"</td></tr>"							
						end if
                        '*************************************
						'LOCALIZAÇÃO (BN)
						'*************************************
						if xmlCampos.nodeName = "LOCALIZACAO" then
							loc = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							loc_desc = xmlCampos.attributes.getNamedItem("Descricao").value
                            ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & loc_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&loc&"</td></tr>"
						end if
                        '*************************************
						'TITULO
						'*************************************
						if xmlCampos.nodeName = "TITULO" then
							titulo = xmlCampos.attributes.getNamedItem("Titulo").value
							titulo_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							complemento = " " & xmlCampos.attributes.getNamedItem("Complemento").value
							titulo_f = Replace(Replace(titulo,"<","&lt;"),">","&gt;")							
							pos_dec = Registro - 1
																			
 							ficha = ficha & "<tr>"
 							ficha = ficha & "<td class='td_imp_esq direita'>"
 							ficha = ficha & titulo_desc
 							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"
 							ficha = ficha & titulo&complemento
 							ficha = ficha & "</td></tr>"	
						end if
						'*************************************
						'DATA
						'*************************************
						if xmlCampos.nodeName = "DATA" then
							data = xmlCampos.attributes.getNamedItem("Valor").value
							data_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & data_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>"&data&"</td></tr>"
						end if
						'*************************************
						'ASSUNTO
						'*************************************
						if xmlCampos.nodeName = "ASSUNTOS" then
							seq = 1
							assunto = ""
							exibe_todos_ass = trim(xmlCampos.attributes.getNamedItem("ExibeTodos").value)
							desc_assunto = trim(xmlCampos.attributes.getNamedItem("Descricao").value)
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'*************************************
								'SE EXIBE TODOS OS ASSUNTOS NA FICHA
								'*************************************
								if (exibe_todos_ass = "1") then
									tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
									desc_princ_assunto = replace(replace(trim(xmlSubCampos.attributes.getNamedItem("Desc_Princ").value)," ","_"),"&#39;","_#39;")
									valor_assunto = trim(xmlSubCampos.attributes.getNamedItem("Valor").value)
									seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)

									temp = seq_assunto&" "&replace(valor_assunto,chr(10),"<br/>")

									if assunto <> "" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & "&nbsp;" & temp & "&nbsp;"
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='autLink td_left_middle'>" & temp & "&nbsp;</td></tr></table>"
										end if
									else
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & temp & "&nbsp;"
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='autLink td_left_middle'>" & temp & "&nbsp;</td></tr></table>"
										end if
									end if
									seq = seq + 1
								'*************************************
								'SE EXIBE SOMENTE PRIMEIRO ASSUNTO
								'*************************************
								else
									desc_assunto = xmlSubCampos.attributes.getNamedItem("Valor").value
									assunto_desc = xmlSubCampos.attributes.getNamedItem("Descricao").value
									assunto_princ = xmlSubCampos.attributes.getNamedItem("Desc_Princ").value
									assunto_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							
									assunto_formatado = replace(replace(replace(assunto_princ," ","_"),"<",""),">","")
																	
									ficha = ficha & "<td class='td_imp_dir esquerda'>"&desc_assunto&"</td></tr>"
								end if
							Next
							if (exibe_todos_ass = "1") then
								ficha = ficha & "<td class='td_imp_dir esquerda'>" & assunto & "</td>"
								ficha = ficha & "</tr>"
							end if
						end if									
						'*************************************
						'FONTES
						'*************************************
						if xmlCampos.nodeName = "FONTES" then
							For Each xmlFontes In xmlCampos.childNodes
								if xmlFontes.nodeName = "FONTE_OBRA" or xmlFontes.nodeName = "FONTE_PER" then
									For Each xmlFonte In xmlFontes.childNodes							
										sDesc_fonte   = xmlFonte.attributes.getNamedItem("DESCRICAO").value
										sTit_fonte    = xmlFonte.attributes.getNamedItem("TITULO").value
										sComp_fonte   = xmlFonte.attributes.getNamedItem("COMPLEMENTO").value
										sPreCom_fonte = xmlFonte.attributes.getNamedItem("PRE_COMP").value
										sTipo_fonte   = xmlFonte.attributes.getNamedItem("TIPO").value
										sCod_fonte    = xmlFonte.attributes.getNamedItem("CODIGO").value
																										
										ficha = ficha & "<tr>"
										ficha = ficha & "<td class='td_imp_esq direita'>"
										ficha = ficha & sDesc_fonte
										ficha = ficha & "</td><td class='td_imp_dir esquerda'>"
										ficha = ficha & sPreCom_fonte&sTit_fonte&sComp_fonte
										ficha = ficha & "</td></tr>"
									Next
								end if
							Next
						end if
						'*************************************
						'SITES
						'*************************************
						if xmlCampos.nodeName = "SITE" then
							site_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							site_c = ""
							For Each xmlItem In xmlCampos.childNodes
								site = TrocaTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
								site_url = RemoveTagMarcador(trim(xmlItem.attributes.getNamedItem("Valor").value))
								if site_c <> "" then
									site_c = site_c + "<br/>"
								end if
								site_c = site_c + "<p>" & site & "</p>"
							Next	

							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_imp_esq direita'>"
							ficha = ficha & site_desc
							ficha = ficha & "</td><td class='td_imp_dir esquerda'>" & site_c & "</td></tr>"
						end if
					end if							
				Next
				
				if tem_chamada = false then
					ficha = Replace(ficha,"colspan=2","")
				end if
				
				Response.Write "<table style='border-spacing: 2px; padding: 1px; margin-bottom: 4px; width: 98%'>"&ficha&"</table>"
				Response.Write "</td>"
			end if
			sequencial = sequencial + 1
		Next
	End if	
	Set xmlDoc = nothing
	Set xmlRoot = nothing		
End if

Response.Write "</table>"
%>
