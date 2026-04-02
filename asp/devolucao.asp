<%
'----------------------------------------------------------------------
'----------------------- D E V O L U Ç Ã O ----------------------------
'----------------------------------------------------------------------
numero_circulacao = trim(Request.QueryString("num_circulacao"))

On Error Resume Next

Set ROService = ROServer.CreateService("Web_Consulta")	
sXMLRenova = ROService.DevolverEmprestimo(numero_circulacao, global_idioma)
Set ROService = nothing

TrataErros(1)

if sXMLRenova <> "" then
	Response.Write "<br />"
	
	sGridRenovacao   = ""
	if (left(sXMLRenova,5) = "<?xml") then
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml sXMLRenova
		Set xmlRoot = xmlDoc.documentElement
		if xmlRoot.nodeName = "RENOVACAO" then
			For Each xmlPNode In xmlRoot.childNodes
				'*************************************
				'INFORMAÇÃO DOS USUÁRIOS
				'*************************************				
				if xmlPNode.nodeName = "INFOUSU" then
					sGridRenovacao = sGridRenovacao & "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
					sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'><td class='td_tabelas_titulo centro' colspan=2><b>"&getTermo(global_idioma, 9004, "Dados da devolução", 0)&"</b><br /></td></tr>"
					For Each xmlUsu In xmlPNode.childNodes
						'*************************************
						'NOME DO USUÁRIO
						'*************************************
						if xmlUsu.nodeName  = "NOME_USU" then	
							Nome_Usuario = xmlUsu.attributes.getNamedItem("VALOR").value
							Desc_Nome_Usuario = xmlUsu.attributes.getNamedItem("DESCRICAO").value					
							sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
							sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width:172px'>&nbsp;"&Desc_Nome_Usuario&"</td>"
							sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&Nome_Usuario&"</td>"
							sGridRenovacao = sGridRenovacao & "</tr>" 
						end if
						'*************************************
						'CODIGO OU MATRICULA DO USUÁRIO
						'*************************************
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
				'*************************************
				'INFORMAÇÃO DAS CIRCULAÇÕES NÃO RENOVADAS
				'*************************************				
				if xmlPNode.nodeName = "NAO_RENOVADOS" then
					if xmlPNode.attributes.getNamedItem("QUANTIDADE").value > 0 then
						iNumReg = 1
						sGridRenovacao = sGridRenovacao & "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
						sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'><td class='td_tabelas_titulo centro' colspan=2><b>"&getTermo(global_idioma, 9005, "Circulações não devolvidas", 0)&"</b><br /></td></tr>"
						sGridRenovacao = sGridRenovacao & "</table>"
						For Each xmlRenova In xmlPNode.childNodes
							sGridRenovacao = sGridRenovacao & "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
							'*************************************
							'INFORMAÇÃO DAS CIRCULAÇÕES NÃO RENOVADAS
							'*************************************				
							if xmlRenova.nodeName = "CIRCULA" then
								For Each xmlCirc In xmlRenova.childNodes
									'*************************************
									'TÍTULO DA OBRA
									'*************************************
									if xmlCirc.nodeName  = "OBRA" then	
										Titulo_Obra = xmlCirc.attributes.getNamedItem("VALOR").value
										Desc_Titulo_Obra = xmlCirc.attributes.getNamedItem("DESCRICAO").value					
										sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 centro' style='width: 20px' rowspan=2><b>" & iNumReg & "</b></td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width: 150px'>&nbsp;"&Desc_Titulo_Obra&"</td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&Titulo_Obra&"</td>"
										sGridRenovacao = sGridRenovacao & "</tr>" 
									end if
									'*************************************
									'MOTIVO
									'*************************************
									if xmlCirc.nodeName  = "MOTIVO" then	
										Motivo = xmlCirc.attributes.getNamedItem("VALOR").value
										Desc_Motivo = xmlCirc.attributes.getNamedItem("DESCRICAO").value
										sGridRenovacao = sGridRenovacao & "<tr style='height: 20px' style='vertical-align: top'>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width: 150px'>&nbsp;"&Desc_Motivo&"</td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;<span style='color: #990000'>"&Motivo&"</span></td>"
										sGridRenovacao = sGridRenovacao & "</tr>" 
									end if
								Next
							end if
							sGridRenovacao = sGridRenovacao & "</table>"
							iNumReg = iNumReg + 1
						Next
					end if
				end if
			Next

            For Each xmlPNode In xmlRoot.childNodes
				'*************************************
				'INFORMAÇÃO DAS CIRCULAÇÕES RENOVADAS
				'*************************************				
				if xmlPNode.nodeName = "RENOVADOS" then
					if xmlPNode.attributes.getNamedItem("QUANTIDADE").value > 0 then
						if (global_multibib = 1) then
                            quantLinhas = 6
                        else
                            quantLinhas = 5    
                        end if
                        
                        iNumReg = 1
						sGridRenovacao = sGridRenovacao & "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
						sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'><td class='td_tabelas_titulo centro' colspan=2><b>"&getTermo(global_idioma, 9007, "Circulações devolvidas", 0)&"</b><br /></td></tr>"
						sGridRenovacao = sGridRenovacao & "</table>"
						For Each xmlRenova In xmlPNode.childNodes
							sGridRenovacao = sGridRenovacao & "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
							'*************************************
							'INFORMAÇÃO DAS CIRCULAÇÕES RENOVADAS
							'*************************************
							if xmlRenova.nodeName = "CIRCULA" then
								For Each xmlCirc In xmlRenova.childNodes
									'*************************************
									'CÓDIGO DA RENOVAÇÃO
									'*************************************
									if xmlCirc.nodeName  = "COD_RENOVA" then	
										Cod_Circ = xmlCirc.attributes.getNamedItem("VALOR").value
										Desc_Cod_Circ = xmlCirc.attributes.getNamedItem("DESCRICAO").value
										sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 centro' style='width: 20px' rowspan=" & quantLinhas &"><b>" & iNumReg & "</b></td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width: 150px'>&nbsp;"&Desc_Cod_Circ&"</td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&Cod_Circ&"</td>"
										sGridRenovacao = sGridRenovacao & "</tr>"

									end if
									'*************************************
									'TÍTULO DA OBRA
									'*************************************
									if xmlCirc.nodeName  = "OBRA" then	
										Titulo_Obra = xmlCirc.attributes.getNamedItem("VALOR").value
										Desc_Titulo_Obra = xmlCirc.attributes.getNamedItem("DESCRICAO").value
										sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width: 150px'>&nbsp;"&Desc_Titulo_Obra&"</td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda' >&nbsp;"&Titulo_Obra&"</td>"
										sGridRenovacao = sGridRenovacao & "</tr>"
										
									end if
                                    '*************************************
									'BIBLIOTECA DO EXEMPLAR
									'*************************************
									if xmlCirc.nodeName  = "BIBLIOTECA_EXEMPLAR" then	
										biblioteca_exemplar = xmlCirc.attributes.getNamedItem("VALOR").value
										desc_biblioteca_exemplar = xmlCirc.attributes.getNamedItem("DESCRICAO").value
                                        codigoBib = xmlCirc.attributes.getNamedItem("CODIGO_BIBLIOTECA").value

                                        descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(codigoBib) & ",0,0);'>"&biblioteca_exemplar&"</a>"
                                        
                                        sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width: 150px'>&nbsp;"&desc_biblioteca_exemplar&"</td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&descricaoBib&"</td>"
										sGridRenovacao = sGridRenovacao & "</tr>" 

									end if

									'*************************************
									'CLASSIFICACAO DA OBRA
									'*************************************
									if xmlCirc.nodeName  = "CLASSIFICACAO" then	
										Classif_Obra = xmlCirc.attributes.getNamedItem("VALOR").value
										Desc_Classif_Obra = xmlCirc.attributes.getNamedItem("DESCRICAO").value
									end if
									'*************************************
									'EXEMPLAR
									'*************************************
									if xmlCirc.nodeName  = "EXEMPLAR" then	
										Exemplar_Obra = xmlCirc.attributes.getNamedItem("VALOR").value
										Desc_Exemplar_Obra = xmlCirc.attributes.getNamedItem("DESCRICAO").value
									end if
									'*************************************
									'DATA DE SAÍDA
									'*************************************
									if xmlCirc.nodeName  = "DATA_SAIDA" then	
										Data_saida = xmlCirc.attributes.getNamedItem("VALOR").value
										Desc_Data_saida = xmlCirc.attributes.getNamedItem("DESCRICAO").value					
										sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width: 150px'>&nbsp;"&Desc_Data_saida&"</td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&Data_saida&"</td>"
										sGridRenovacao = sGridRenovacao & "</tr>" 
										
									end if
									'*************************************
									'DATA PREVISTA
									'*************************************
									if xmlCirc.nodeName  = "DATA_DEV" then	
										data_prevista = xmlCirc.attributes.getNamedItem("VALOR").value
										desc_data_prevista = xmlCirc.attributes.getNamedItem("DESCRICAO").value
										sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width: 150px'>&nbsp;"&desc_data_prevista&"</td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;"&data_prevista&"</td>"
										sGridRenovacao = sGridRenovacao & "</tr>"
										
									end if
									'*************************************
									'RENOVAÇÕES PERMITIDAS
									'*************************************
									if xmlCirc.nodeName  = "REN_PERMITIDAS" then
										ren_permitidas = xmlCirc.attributes.getNamedItem("VALOR").value									
										desc_ren_permitidas = xmlCirc.attributes.getNamedItem("DESCRICAO").value
										sGridRenovacao = sGridRenovacao & "<tr style='height: 20px'>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor1 esquerda' style='width: 150px'>&nbsp;"&desc_ren_permitidas&"</td>"
										sGridRenovacao = sGridRenovacao & "<td class='td_tabelas_valor2 esquerda'>&nbsp;<span style='color: blue'>"&ren_permitidas&"</span></td>"
										sGridRenovacao = sGridRenovacao & "</tr>"
									end if
								Next
							end if
							sGridRenovacao = sGridRenovacao & "</table>"
							
							iNumReg = iNumReg + 1
						Next
					end if
				end if
			Next
		end if
		Response.Write "<br />"
		sLinkVoltar = "<span class='transparent-icon span_imagem icon_16 icon-small-back'></span>&nbsp;<a class='link_serv' href='javascript:LinkDevolucao(parent.hiddenFrame.modo_busca);'>"&getTermo(global_idioma, 4096, "Voltar para a tela de circulações", 0)&"</a>"
		Response.write "<table style='width: 100%'>"
		Response.write "<tr style='height: 25px'>"
		Response.write "<td class='td_servicos_titulo td_left_middle'>"
		Response.Write "<table style='width: 100%'><tr><td class='td_left_middle'>"
		Response.write "&nbsp;"&sLinkVoltar&"&nbsp;"
		Response.Write "</td></tr></table>"
		Response.write "</td></tr></table>"
		Response.Write sGridRenovacao
		Response.Write "<br />"
	end if
end if
%>