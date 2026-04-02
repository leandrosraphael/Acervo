<%	
'--------------------------------------------------------------------------------
' 				MONTA RESULTADO EM FORMATO DE FICHA
'--------------------------------------------------------------------------------
Response.Write "<table class='max_width' style='border-spacing: 1px; padding: 0'><tr>"
max_registros = 1000
'**************************************************************************	
' ADEQUAÇÃO UNICAMP: Exibir no máximo 500 resultados
' ADEQUAÇÃO BNBD: Exibir no máximo 10000 resultados
'**************************************************************************
if (global_numero_serie = 5516) then
	max_registros = 10000
elseif (global_numero_serie = 5592) then
    max_registros = 200
else
    max_registros = 1000
end if

if GetSessao(global_cookie,"nrows_real"&iIndexSrv) <> "" then
	if ((CLng(GetSessao(global_cookie,"nrows_real"&iIndexSrv)) > max_registros) and (sOrigem <> "MySel")) then
		num_total = GetSessao(global_cookie,"nrows"&iIndexSrv)
		qtd_links = int(num_total/global_num_linhas)
		if (num_total mod global_num_linhas) > 0 then
			qtd_links = qtd_links + 1
		end if
		msg_busca = getTermo(global_idioma, 2637, "Apresentando os primeiros 1000 registros em %s páginas.", 0)
		msg_busca = replace(msg_busca, "1000", CStr(max_registros))
		msg_busca = Format(msg_busca, qtd_links) & " " & getTermo(global_idioma, 2667, "Por favor, refine sua busca.", 0)

		if global_exibe_capa = 1 then
			sColFicha = "4"
		else
			sColFicha = "3"
		end if

		Response.Write "<td class='td_grid_ficha_background td_grid_ficha_borda td_grid_ficha_total_resultados centro' style='border-bottom: 1px solid #ccc;' colspan='" & sColFicha & "'>"
		Response.Write "<span class='span_imagem div_imagem_right icon_16_15 alerta0'></span>&nbsp;" & msg_busca 
		Response.Write "</td></tr><tr>"
	end if
end if

if (left(sXMLFichas,5) = "<?xml") then
	sequencial = 1
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml sXMLFichas
	Set xmlRoot = xmlDoc.documentElement      
    
	if xmlRoot.nodeName = "Pagina" then
        
		For Each xmlFicha In xmlRoot.childNodes
            if xmlFicha.nodeName  = "IDENTIFICADORES" then                
                For Each xmlIdentificadores in xmlFicha.childNodes
                    tipoIdentificador = xmlIdentificadores.attributes.getNamedItem("Tipo").value
                    imagemIdentificador = "data:image;base64," & xmlIdentificadores.attributes.getNamedItem("Base64").value
                    %><script>sessionStorage.setItem("<%=tipoIdentificador%>" ,"<%=imagemIdentificador%>");</script><%                    
                next
            end if
			if xmlFicha.nodeName  = "Ficha" then
				Tipo = xmlFicha.attributes.getNamedItem("Tipo").value
				Registro = xmlFicha.attributes.getNamedItem("Registro").value
				
				link_class = "link_serv"
				Response.Write "<tr>"
				'//-------------------------------------------------------------//
				'---------------------CELULA COM IMAGEM E NUMERO ---------------//
				'//-------------------------------------------------------------//
				if cStr(Tipo) = "0" then
					css = "icon-paper"
				elseif cStr(Tipo) = "1" then
					css = "icon-book"
				elseif cStr(Tipo) = "2" then
					css = "icon-rules"
				elseif cStr(Tipo) = "3" then
					css = "icon-document"
				end if

				if (sOrigem <> "emprestimo") then
					Response.Write "<td class='td_center_top td_grid_ficha_background' style='width:40px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"

					Response.Write "<br /><table style='display: inline-block'><tr><td class='centro' style='font-size: 13;'>"
					Response.Write Registro&"<br />"
					Response.Write "<div class='transparent-icon icon_24 "&css&"'>&nbsp;</div>"
					Response.Write "</td></tr></table>"
						
					Response.Write "</td>"				
				end if
				'//-------------------------------------------------------------//
				'---------------------CELULA COM A FICHA------------------------//
				'//-------------------------------------------------------------//
				ficha = ""
				codigo_atual = 0
				tem_chamada = false
				
				'*************************************
				'Serviços
				'*************************************
				serv_selecao = false
				serv_reserva = false
				serv_aquisicao = false
				serv_midias = false
				serv_exemplar = false
				serv_analitica = false
                serv_referencia = false

				'*************************************
				'EXTRAÇÃO DE TÍTULO E AUTOR, QUANDO NÃO EXISTIR CAPA PARA O ELEMENTO PESQUISADO
				'*************************************
				codigo_material = ""
				titulo_sem_capa = ""
				desc_autor_sem_capa = ""
				url_imagem_referencia = ""
				titulo_autor_sem_capa = "<div class='div_capa'>"
				For Each xmlCampos In xmlFicha.childNodes
					if xmlCampos.nodeName = "TITULO" OR xmlCampos.nodeName = "NORMA" then
						if (xmlCampos.nodeName = "TITULO") then
							titulo_sem_capa = xmlCampos.attributes.getNamedItem("Titulo").value
							titulo_f_sem_capa = Replace(Replace(titulo_sem_capa,"#D#",""),"#/D#","")
							complemento_sem_capa = " " & Replace(Replace(xmlCampos.attributes.getNamedItem("Complemento").value,"#D#",""),"#/D#","")
							titulo_f_sem_capa = titulo_f_sem_capa & complemento_sem_capa
							if( Len(titulo_f_sem_capa) >= 45) then
								titulo_f_sem_capa = left(titulo_f_sem_capa,45) & "..."
							end if    
						elseif xmlCampos.nodeName = "NORMA" then
							titulo_sem_capa = xmlCampos.attributes.getNamedItem("Valor").value
							titulo_f_sem_capa = Replace(Replace(titulo_sem_capa,"#D#",""),"#/D#","")
							titulo_f_sem_capa = titulo_f_sem_capa & complemento_sem_capa
							if (Len(titulo_f_sem_capa) >= 45) then
								titulo_f_sem_capa = left(titulo_f_sem_capa,45) & "..."
							end if
						end if       					    
					elseif xmlCampos.nodeName = "ENT_PRINC" OR xmlCampos.nodeName = "ORGAO_ORIGEM" then
						desc_autor_sem_capa = xmlCampos.attributes.getNamedItem("Valor").value 
						desc_autor_sem_capa = Replace(Replace(desc_autor_sem_capa,"#D#",""),"#/D#","")
						if( Len(desc_autor_sem_capa) >= 25 ) then
							desc_autor_sem_capa = left(desc_autor_sem_capa,25) & "..."
						end if
						desc_autor_sem_capa = "<br/><br/>" & "<span>" & desc_autor_sem_capa & "</span>"
					' Adequação BND - A imagem de referência irá remeter para o site cadastrado
					elseif xmlCampos.nodeName = "SITE" AND global_numero_serie = 5516 AND url_imagem_referencia = "" then
                        For Each xmlSubCampos In xmlCampos.childNodes
                            codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
                            dspace = xmlSubCampos.attributes.getNamedItem("DSpace").value
						    url_site = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
						    if (UCase(Right(url_site, 3)) = "JPG") then
							    url_imagem_referencia = url_site
                                Exit For
						    end if
                        Next
					elseif xmlCampos.nodeName = "MIDIA" AND (global_numero_serie = 5516 or global_numero_serie = 5631) then
						codigo_material = xmlCampos.attributes.getNamedItem("Codigo").value
					end if
				next

				' Adequação BND - A imagem de referência irá remeter para o site cadastrado (PDF, MP3, MID)
				if global_numero_serie = 5516 AND url_imagem_referencia = "" then
					For Each xmlCampos In xmlFicha.childNodes
						if xmlCampos.nodeName = "SITE" AND url_imagem_referencia = "" then
                            For Each xmlSubCampos In xmlCampos.childNodes
                                codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
                                dspace = xmlSubCampos.attributes.getNamedItem("DSpace").value
							    url_site = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
							    if (UCase(Right(url_site, 3)) = "PDF") then
								    url_imagem_referencia = url_site
                                    Exit For
							    end if
                            Next
						end if
					Next
					if url_imagem_referencia = "" then
						For Each xmlCampos In xmlFicha.childNodes
							if xmlCampos.nodeName = "SITE" AND url_imagem_referencia = "" then
                                For Each xmlSubCampos In xmlCampos.childNodes
                                    codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
                                    dspace = xmlSubCampos.attributes.getNamedItem("DSpace").value
								    url_site = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								    if (UCase(Right(url_site, 3)) = "MP3") then
									    url_imagem_referencia = url_site
									    capa_audio = true
                                        Exit For
								    end if
                                Next
							end if
						Next
					end if
					if url_imagem_referencia = "" then
						For Each xmlCampos In xmlFicha.childNodes
							if xmlCampos.nodeName = "SITE" AND url_imagem_referencia = "" then
                                For Each xmlSubCampos In xmlCampos.childNodes
                                    codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
                                    dspace = xmlSubCampos.attributes.getNamedItem("DSpace").value
								    url_site = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								    if (UCase(Right(url_site, 3)) = "MID") then
									    capa_audio = true
                                        Exit For
								    end if
                                Next
							end if
						Next
					end if
					if url_imagem_referencia = "" then
						For Each xmlCampos In xmlFicha.childNodes
							if xmlCampos.nodeName = "SITE" AND url_imagem_referencia = "" then
                                For Each xmlSubCampos In xmlCampos.childNodes
                                    codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
                                    dspace = xmlSubCampos.attributes.getNamedItem("DSpace").value                                    
								    url_site = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								    url_imagem_referencia = url_site
                                    Exit For
                                Next
							end if
						Next
					end if
				end if

				titulo_autor_sem_capa = titulo_autor_sem_capa & titulo_f_sem_capa & desc_autor_sem_capa & "</div>"

				if (url_imagem_referencia <> "") then
					titulo_autor_sem_capa = titulo_autor_sem_capa & "<div class='mascara_capa_fantasia_transparente' data-url=""" & url_imagem_referencia & """ onclick=""abrirLink(this);ContarAcesso(#COD#,"&codMidia&",1,"&dspace&");""></div>"
				end if
				'*************************************
				'FIM EXTRAÇÃO DE TÍTULO E AUTOR, QUANDO NÃO EXISTIR CAPA PARA O ELEMENTO PESQUISADO
				'*************************************
				
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
							if titulo_autor_sem_capa <> "" then
								titulo_autor_sem_capa = Replace(titulo_autor_sem_capa, "#COD#", codigo_atual)
							end if
							
							if global_exibe_capa = 1 then
								if CStr(xmlCampos.attributes.getNamedItem("Capa").value) = "1" then
									Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
									Response.Write "<table style='display: inline-table'><tr><td>"
									Response.Write "<br />"
									if url_imagem_referencia <> "" then
										Response.Write "<a href='" & url_imagem_referencia & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>"
									end if
									Response.Write "<img alt='' src='asp/capa.asp?obra="&codigo_atual&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'>"
									if url_imagem_referencia <> "" then
										Response.Write "</a>"
									end if
									Response.Write "<br />&nbsp;"
									Response.Write "</td></tr></table>"
									Response.Write "</td>"
								else							
									if (codigo_material = "46" or codigo_material = "35") then
										Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
										Response.Write "<table style='display: inline-table'><tr><td>"
										Response.Write "<br />"
										if url_imagem_referencia <> "" then
											Response.Write "<a href='" & url_imagem_referencia & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>"
										end if
										Response.Write "<div class='capa_audio icon_capa'></div>"
										if url_imagem_referencia <> "" then
											Response.Write "</a>"
										end if
										Response.Write "<br />&nbsp;"
										Response.Write "</td></tr></table>"
										Response.Write "</td>"
									else					
										Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
										Response.Write "<table style='display: inline-table'><tr><td>"
										Response.Write "<br/>"
										Response.Write  titulo_autor_sem_capa
										Response.Write "<div class='capa_fantasia icon_capa'></div>"
										Response.Write "<br/>"
										Response.Write "</td></tr></table>"
										Response.Write "</td>"
									end if
								end if
							end if
							Response.Write "<td class='td_center_top td_grid_ficha_background' style='border-bottom: 1px solid #ccc;'>"
						end if
						'*************************************
						'MATERIAL
						'*************************************
						if xmlCampos.nodeName = "MIDIA" then
							material = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							material_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & material_desc
							ficha = ficha & "</td><td class='esquerda td_ficha_dir'>"&material&"</td></tr>"
						end if
						'*************************************
						'NUMERO DE CHAMADA
						'*************************************
						if xmlCampos.nodeName = "CHAMADA" then
							chamada = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							chamada_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & chamada_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&chamada&"</td></tr>"
						end if
						'*************************************
						'ANDAR (UNICENP)
						'*************************************
						if xmlCampos.nodeName = "ANDAR" then
							andar = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							andar_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & andar_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&andar&"</td></tr>"
						end if
                        '*************************************
						'LOCALIZAÇÃO (BN)
						'*************************************
						if xmlCampos.nodeName = "LOCALIZACAO" then
							loc = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							loc_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & loc_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&loc&"</td></tr>"
						end if
						'*************************************
						'TITULO
						'*************************************
						if xmlCampos.nodeName = "TITULO" then
							titulo = xmlCampos.attributes.getNamedItem("Titulo").value
							titulo_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							complemento = " " & TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Complemento").value)
							titulo_f = Replace(Replace(titulo,"<","&lt;"),">","&gt;")							
							pos_dec = Registro - 1
							if(sOrigem = "Pesquisa") then
								total = GetSessao(global_cookie,"nrows"&iIndexSrv)
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&","&pagina&",'resultado',"&Tipo&");"">"&titulo_f &"</a>"
							elseif(sOrigem = "MySel") then
								total = iNumSel
								titulo = "<a class='link_custom_negrito' title="&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'sels',"&Tipo&");"">"&titulo_f &"</a>"
							elseif(sOrigem = "BibCurso") then
								total = Resultado
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'bibcurso',"&Tipo&");"">"&titulo_f &"</a>"
							elseif(sOrigem = "favoritos") then
								total = quantidade
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'favoritos',"&Tipo&");"">"&titulo_f &"</a>"
							end if
							
							titulo = TrocaTagMarcador(titulo)
							
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & titulo_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"
							ficha = ficha & titulo&complemento
							ficha = ficha & "</td></tr>"	
						end if
						'*************************************
						'OUTROS TITULOS (Adequação UNIUBE)
						'*************************************
						if xmlCampos.nodeName = "OUTROS_TITULOS" then
							For Each xmlSubCampos In xmlCampos.childNodes
								desc_outros_titulos = trim(xmlSubCampos.attributes.getNamedItem("Descricao").value)
								valor_outros_titulos = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								
								ficha = ficha & "<tr>"
								ficha = ficha & "<td class='td_ficha_esq direita'>"
								ficha = ficha & desc_outros_titulos
								ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&valor_outros_titulos&"</td></tr>"
							Next
						end if 
						'*************************************
						'ANO
						'Adequeação IMS: não mostrar
						'*************************************
						if xmlCampos.nodeName = "ANO" and (global_numero_serie <> 4516) then
							ano = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							ano_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & ano_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&ano&"</td></tr>"
						end if
						'************************************************************
						' IMPRENTA
						'Adequeação: apenas IMS
						'************************************************************
						if xmlCampos.nodeName = "IMPRENTA" and (global_numero_serie = 4516) then
							imprenta = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_imprenta = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & desc_imprenta
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&imprenta&"</td></tr>"
						end if
						'************************************************************
						' Nº ÁLBUM
						'Adequeação: apenas IMS
						'************************************************************
						if xmlCampos.nodeName = "COMPLEMENTO1" and (global_numero_serie = 4516) then
							complemento = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_complemento = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & desc_complemento
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&complemento&"</td></tr>"
						end if
						'*************************************
						'PERIODICIDADE
						'*************************************
						if xmlCampos.nodeName = "PERIODICIDADE" then
							periodicidade = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							periodicidade_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & periodicidade_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&periodicidade&"</td></tr>"
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
							ficha = ficha & "<td class='td_ficha_esq direita'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'*************************************
								'SE EXIBE TODOS OS ASSUNTOS NA FICHA
								'*************************************
								if (exibe_todos_ass = "1") then
									tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
									valor_assunto = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
									busca_assunto = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
									desc_princ_assunto = replace(replace(replace(busca_assunto," ","_"),"&#39;","_#39;"),"'","\'")
									seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)
									nome_funcao = "LinkBuscaAssunto"
									hint = getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"..."
									if (sOrigem <> "emprestimo") and xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assInfo = "<span class='assuntoImg span_imagem icon_16' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</span>"
										else
											assInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
										end if
									else
										assInfo = ""
									end if
									temp = seq_assunto&" <a class='link_classic2' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_assunto&",'"&desc_princ_assunto&"',"&iIndexSrv&");"">"&replace(valor_assunto,chr(10),"<br />") &"</a>"
                                    identificadoresInfo = ""
                                    For Each xmlIdentificadores in xmlCampos.childNodes                                
	                                    %><!--#include file ="identificadoresDetalhe.asp"--><%
                                    next
                                    assInfo = assInfo & identificadoresInfo
									if assunto <> "" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & "&nbsp;" & temp & "&nbsp;" & assInfo
										else
											assunto = assunto & "<table class='autLink td_left_middle' style='border-spacing: 1px ; padding: 0'><tr><td id='assLinkTab"&seq&"'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
										end if
									else
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & temp & "&nbsp;" & assInfo
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px ; padding: 0'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
										end if
									end if
									seq = seq + 1
								'*************************************
								'SE EXIBE SOMENTE PRIMEIRO ASSUNTO
								'*************************************
								else
									desc_assunto = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
                                    tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									assunto_desc = xmlSubCampos.attributes.getNamedItem("Descricao").value
									assunto_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value
									assunto_busca = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
									assunto_formatado = replace(replace(replace(replace(assunto_busca," ","_"),"<",""),">",""),"'","\'")
									if (sOrigem <> "emprestimo") and xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
										assuntoInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&sequencial&"' onClick=LinkAutInfo('linkAutInfo',"&assunto_codigo&",'"&assunto_formatado&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
									else
										assuntoInfo = ""
									end if

                                    identificadoresInfo = ""
                                    For Each xmlAssuntos in xmlCampos.childNodes                                
                                        For Each xmlIdentificadores in xmlAssuntos.childNodes
                                            if xmlIdentificadores.nodename = "IDENTIFICADOR" then
	                                        %><!--#include file ="identificadoresDetalhe.asp"--><%
                                            end if
                                        next
                                    next
                                    assuntoInfo = assuntoInfo & identificadoresInfo
									assunto = "<table class='autLink remover_bordas_padding'><tr><td id='assLinkTab"&sequencial&"' class='td_left_middle'><a class='link_custom' title='"&getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"...' href=""javascript:LinkBuscaAssunto(parent.hiddenFrame.modo_busca,"&assunto_codigo&",'"&assunto_formatado&"',"&iIndexSrv&");"">"&desc_assunto&"</a>&nbsp;</td>"&assuntoInfo&"</tr></table>"
									ficha = ficha & "<td class='td_ficha_dir esquerda'>"&assunto&"</td></tr>"
									if len(trim(desc_assunto)) > 85 then
										ficha = ficha &"<script>redimensiona('assLinkTab"&sequencial&"',476);</script>"
									end if
								end if
							Next
							if (exibe_todos_ass = "1") then
								ficha = ficha & "<td class='td_ficha_dir esquerda'>" & assunto & "</td>"
								ficha = ficha & "</tr>"
							end if						
						end if
						'*************************************
						'SITES RELACIONADOS
						'*************************************
						if xmlCampos.nodeName = "SITE" then
							site_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							site_c = ""
							For Each xmlSubCampos In xmlCampos.childNodes
								codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
								plataforma = xmlSubCampos.attributes.getNamedItem("Plataforma").value
								site = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								site_url = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
                                dspace = xmlSubCampos.attributes.getNamedItem("DSpace").value
								if site_c <> "" then
									site_c = site_c & "<br/>"
								end if
								if (plataforma <> "") then
									site_c = site_c & "<a class='link_classic2' href='javascript:abreMidiaEspecifica(" & Tipo & "," & codigo_atual & "," & codMidia & "," & iIndexSrv & ")'>" & site & "</a>"
								else
									site_c = site_c & "<a class='link_classic2' href='" & site_url & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>" & site & "</a>"
								end if
							Next	
														
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & site_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda' style='word-break: break-all;'>" & site_c & "</td></tr>"
						end if

                        '*************************************
						'ACERVO - QUANTIDADE DE EXEMPLARES
						'*************************************
                        if xmlCampos.nodeName = "ACERVO" then
                            acervo_qtdes = ""
                            acervo_desc = xmlCampos.attributes.getNamedItem("Descricao").value 
                            ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"&acervo_desc&"</td>"
                            ficha = ficha & "<td class='esquerda td_ficha_dir'>"
                            For Each xmlSubCampos In xmlCampos.childNodes
                                if acervo_qtdes <> "" then
                                    ficha = ficha & "<br/><br/>"    
                                end if
                                acervo_qtdes = xmlSubCampos.attributes.getNamedItem("Descricao").value
                                ficha = ficha & acervo_qtdes
                            Next	
                            ficha = ficha & "</td></tr>"    
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
							if titulo_autor_sem_capa <> "" then
								titulo_autor_sem_capa = Replace(titulo_autor_sem_capa, "#COD#", codigo_atual)
							end if
							
							if global_exibe_capa = 1 then
								if CStr(xmlCampos.attributes.getNamedItem("Capa").value) = "1" then
									Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
									Response.Write "<table style='display: inline-table'><tr><td>"
									Response.Write "<br />"
									if url_imagem_referencia <> "" then
										Response.Write "<a href='" & url_imagem_referencia & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>"
									end if
									Response.Write "<img alt='' src='asp/capa.asp?obra="&codigo_atual&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'>"
									if url_imagem_referencia <> "" then
										Response.Write "</a>"
									end if
									Response.Write "<br />&nbsp;"
									Response.Write "</td></tr></table>"
									Response.Write "</td>"
								else
									if (codigo_material = "46" or codigo_material = "35") then
										Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
										Response.Write "<table style='display: inline-table'><tr><td>"
										Response.Write "<br />"
										if url_imagem_referencia <> "" then
											Response.Write "<a href='" & url_imagem_referencia & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>"
										end if
										Response.Write "<div class='capa_audio icon_capa'></div>" 
										if url_imagem_referencia <> "" then
											Response.Write "</a>"
										end if
										Response.Write "<br />&nbsp;"
										Response.Write "</td></tr></table>"
										Response.Write "</td>"
									else
										Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
										Response.Write "<table style='display: inline-table'><tr><td>"
										Response.Write "<br/>"
										Response.Write  titulo_autor_sem_capa
										Response.Write "<div class='capa_fantasia icon_capa'></div>"
										Response.Write "<br/>"
										Response.Write "</td></tr></table>"
										Response.Write "</td>"
									end if
								end if
							end if
							Response.Write "<td class='td_center_top td_grid_ficha_background' style='border-bottom: 1px solid #ccc;'>"
						end if
						'*************************************
						'MATERIAL
						'*************************************
						if xmlCampos.nodeName = "MIDIA" then
							material = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							material_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & material_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&material&"</td></tr>"
						end if
						'*************************************
						'NUMERO DE CHAMADA
						'*************************************
						if xmlCampos.nodeName = "CHAMADA" then
							chamada = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							chamada_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & chamada_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&chamada&"</td></tr>"
						end if
						'*************************************
						'ANDAR (UNICENP)
						'*************************************
						if xmlCampos.nodeName = "ANDAR" then
							andar = xmlCampos.attributes.getNamedItem("Valor").value
							andar_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & andar_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&andar&"</td></tr>"
						end if
                        '*************************************
						'LOCALIZACAO (BN)
						'*************************************
						if xmlCampos.nodeName = "LOCALIZACAO" then
							loc = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							loc_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & loc_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&loc&"</td></tr>"
						end if
						'*************************************
						'AUTOR (ENTRADA PRINCIPAL)
						'*************************************
						if xmlCampos.nodeName = "ENT_PRINC" then
							desc_autor = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							autor_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							autor_busca = RemoveTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							autor_tipo = xmlCampos.attributes.getNamedItem("Tipo").value
							autor_codigo = xmlCampos.attributes.getNamedItem("Codigo").value							
							nome_funcao = "LinkBuscaAutor"
							autor_formatado = replace(replace(replace(replace(autor_busca," ","_"),"<",""),">",""),"'","\'")
							if (sOrigem <> "emprestimo") and xmlCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
								autInfo = "<td class='autImg'><div class='autImg' id='autImg"&sequencial&"' onClick=LinkAutInfo('linkAutInfo',"&autor_codigo&",'"&autor_formatado&"',"&autor_tipo&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1571, "Mostrar informações sobre este autor", 0)&"...'>&nbsp;</div></td>"
							else
								autInfo = ""
							end if
                            identificadoresInfo = ""
                            For Each xmlIdentificadores in xmlCampos.childNodes                                
	                            %><!--#include file ="identificadoresDetalhe.asp"--><%
                            next
                            autInfo = autInfo & identificadoresInfo
							autor = "<table class='autLink remover_bordas_padding'><tr><td id='autLinkTab"&sequencial&"' class='td_left_middle'><a class='link_custom' title='"&getTermo(global_idioma, 1558, "Buscar todos os registros deste autor", 0)&"...' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&autor_codigo&",'"&autor_formatado&"',"&iIndexSrv&");"">"&desc_autor&"</a>&nbsp;</td>"&autInfo&"</tr></table>"
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & autor_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&autor&"</td></tr>"
                            
							if len(trim(desc_autor)) > 85 then
								ficha = ficha &"<script>redimensiona('autLinkTab"&sequencial&"',476);</script>"
							end if							
						end if
						'*************************************
						'AUTORIA (ENTRADA PRINCIPAL e ENTRADA(S) SECUNDÁRIA(S))
						'Adequação para o cliente IMS
						'*************************************
						if xmlCampos.nodeName = "AUTORIA" then
							autor = ""
							autor_desc = replace(replace(replace(trim(xmlCampos.attributes.getNamedItem("Desc_Autoria").value)," ","_"),"&#39;","_#39;"),"'","\'")
							sequencialOld = sequencial
							For Each xmlSubCampos In xmlCampos.childNodes
								sequencial = sequencial + 1
								desc_autor = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								autor_busca = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
								autor_tipo = xmlSubCampos.attributes.getNamedItem("Tipo").value
								autor_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							
									
								nome_funcao = "LinkBuscaAutor"
								
								autor_formatado = replace(replace(replace(replace(autor_busca," ","_"),"<",""),">",""),"'","\'")
								
								if (sOrigem <> "emprestimo") and xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
									autInfo = "<td class='autImg'><div class='autImg' id='autImg"&sequencial&"' onClick=LinkAutInfo('linkAutInfo',"&autor_codigo&",'"&autor_formatado&"',"&autor_tipo&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1571, "Mostrar informações sobre este autor", 0)&"...'>&nbsp;</div></td>"
								else
									autInfo = ""
								end if
								identificadoresInfo = ""
                                For Each xmlIdentificadores in xmlCampos.childNodes                                
	                                %><!--#include file ="identificadoresDetalhe.asp"--><%
                                next
                                autInfo = autInfo & identificadoresInfo	
								autor = autor & "<table class='autLink remover_bordas_padding'><tr><td id='autLinkTab"&sequencial&"' class='td_left_middle'><a class='link_custom' title='"&getTermo(global_idioma, 1558, "Buscar todos os registros deste autor", 0)&"...' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&autor_codigo&",'"&autor_formatado&"',"&iIndexSrv&");"">"&desc_autor&"</a>&nbsp;</td>"&autInfo&"</tr></table>"
							Next
							sequencial = sequencialOld
								
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & autor_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&autor&"</td></tr>"
							if len(trim(desc_autor)) > 85 then
								ficha = ficha &"<script>redimensiona('autLinkTab"&sequencial&"',476);</script>"
							end if							
						end if
						'*************************************
						'TITULO
						'*************************************
						if xmlCampos.nodeName = "TITULO" then
							titulo = xmlCampos.attributes.getNamedItem("Titulo").value
							titulo_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							complemento = " " & TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Complemento").value)
							titulo_f = Replace(Replace(titulo,"<","&lt;"),">","&gt;")							
							pos_dec = Registro - 1
							
							if(sOrigem = "Pesquisa") then
								total = GetSessao(global_cookie,"nrows"&iIndexSrv)
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&","&pagina&",'resultado',"&Tipo&");"">"&titulo_f &"</a>"
							elseif(sOrigem = "MySel") then
								total = iNumSel
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'sels',"&Tipo&");"">"&titulo_f &"</a>"
							elseif(sOrigem = "BibCurso") then
								total = Resultado
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'bibcurso',"&Tipo&");"">"&titulo_f &"</a>"
							elseif(sOrigem = "favoritos") then
								total = quantidade
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'favoritos',"&Tipo&");"">"&titulo_f &"</a>"
							end if
														
							titulo = TrocaTagMarcador(titulo)
														
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & titulo_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"
							ficha = ficha & titulo&complemento
							ficha = ficha & "</td></tr>"	
						end if
						'*************************************
						'ANO
						'Adequeação IMS: não mostrar
						'*************************************
						if xmlCampos.nodeName = "ANO" and (global_numero_serie <> 4516) then
							ano = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							ano_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & ano_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&ano&"</td></tr>"
						end if
						'************************************************************
						' IMPRENTA
						'Adequeação: apenas IMS
						'************************************************************
						if xmlCampos.nodeName = "IMPRENTA" and (global_numero_serie = 4516) then
							imprenta = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_imprenta = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & desc_imprenta
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&imprenta&"</td></tr>"
						end if
						'************************************************************
						' Nº ÁLBUM
						'Adequeação: apenas IMS
						'************************************************************
						if xmlCampos.nodeName = "COMPLEMENTO1" and (global_numero_serie = 4516) then
							complemento = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_complemento = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & desc_complemento
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&complemento&"</td></tr>"
						end if
						'*************************************
						'DATA
						'*************************************
						if xmlCampos.nodeName = "DATA" then
							data = xmlCampos.attributes.getNamedItem("Valor").value
							data_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & data_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&data&"</td></tr>"
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
							ficha = ficha & "<td class='td_ficha_esq direita'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'*************************************
								'SE EXIBE TODOS OS ASSUNTOS NA FICHA
								'*************************************
								if (exibe_todos_ass = "1") then
									tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
									valor_assunto = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
									busca_assunto = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
									desc_princ_assunto = replace(replace(replace(trim(busca_assunto)," ","_"),"&#39;","_#39;"),"'","\'")
									seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)
									nome_funcao = "LinkBuscaAssunto"
									hint = getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"..."
									if (sOrigem <> "emprestimo") and xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assInfo = "<span class='assuntoImg span_imagem icon_16' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</span>"
										else
											assInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
										end if
									else
										assInfo = ""
									end if
									temp = seq_assunto&" <a class='link_classic2' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_assunto&",'"&desc_princ_assunto&"',"&iIndexSrv&");"">"&replace(valor_assunto,chr(10),"<br />") &"</a>"
                                    identificadoresInfo = ""
                                    For Each xmlAssuntos in xmlCampos.childNodes                                
                                        For Each xmlIdentificadores in xmlAssuntos.childNodes
                                            if xmlIdentificadores.nodename = "IDENTIFICADOR" then
	                                        %><!--#include file ="identificadoresDetalhe.asp"--><%
                                            end if
                                        next
                                    next                                    
                                    assInfo = assInfo & identificadoresInfo
									if assunto <> "" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & "&nbsp;" & temp & "&nbsp;" & assInfo
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
										end if
									else
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & temp & "&nbsp;" & assInfo
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
										end if
									end if
									seq = seq + 1
								'*************************************
								'SE EXIBE SOMENTE PRIMEIRO ASSUNTO
								'*************************************
								else
									desc_assunto = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
                                    tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									assunto_desc = xmlSubCampos.attributes.getNamedItem("Descricao").value
									assunto_busca = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
									assunto_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							
									assunto_formatado = replace(replace(replace(replace(assunto_busca," ","_"),"<",""),">",""),"'","\'")
									if (sOrigem <> "emprestimo") and xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1" then
										assuntoInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&sequencial&"' class='td_left_middle' onClick=LinkAutInfo('linkAutInfo',"&assunto_codigo&",'"&assunto_formatado&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
									else
										assuntoInfo = ""
									end if
                                    identificadoresInfo = ""
                                    For Each xmlAssuntos in xmlCampos.childNodes                                
                                        For Each xmlIdentificadores in xmlAssuntos.childNodes
                                            if xmlIdentificadores.nodename = "IDENTIFICADOR" then
	                                        %><!--#include file ="identificadoresDetalhe.asp"--><%
                                            end if
                                        next
                                    next
                                    assuntoInfo = assuntoInfo & identificadoresInfo
									assunto = "<table class='autLink remover_bordas_padding'><tr><td id='assLinkTab"&sequencial&"' class='td_left_middle'><a class='link_custom' title='"&getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"...' href=""javascript:LinkBuscaAssunto(parent.hiddenFrame.modo_busca,"&assunto_codigo&",'"&assunto_formatado&"',"&iIndexSrv&");"">"&desc_assunto&"</a>&nbsp;</td>"&assuntoInfo&"</tr></table>"
									ficha = ficha & "<td class='td_ficha_dir esquerda'>"&assunto&"</td></tr>"
									if len(trim(desc_assunto)) > 85 then
										ficha = ficha &"<script>redimensiona('assLinkTab"&sequencial&"',476);</script>"
									end if
								end if
							Next
							if (exibe_todos_ass = "1") then
								ficha = ficha & "<td class='td_ficha_dir esquerda'>" & assunto & "</td>"
								ficha = ficha & "</tr>"
							end if						
						end if
						'*************************************
						'SITES
						'*************************************
						if xmlCampos.nodeName = "SITE" then
							site_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							site_c = ""
							For Each xmlSubCampos In xmlCampos.childNodes
								codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
								plataforma = xmlSubCampos.attributes.getNamedItem("Plataforma").value
								site = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								site_url = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
                                dspace = xmlSubCampos.attributes.getNamedItem("DSpace").value
								if site_c <> "" then
									site_c = site_c & "<br/>"
								end if
								if (plataforma <> "") then
									site_c = site_c & "<a class='link_classic2' href='javascript:abreMidiaEspecifica(" & Tipo & "," & codigo_atual & "," & codMidia & "," & iIndexSrv & ")'>" & site & "</a>"
								else
									site_c = site_c & "<a class='link_classic2' href='" & site_url & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>" & site & "</a>"
								end if
							Next	
														
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & site_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda' style='word-break: break-all;'>" & site_c & "</td></tr>"
						end if

						'*************************************
						'SÉRIE - No XML somente quando Espanha
						'*************************************
						if xmlCampos.nodeName = "SERIE" then
							serie = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_serie = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & desc_serie
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&serie&"</td></tr>"
						end if

                        '*************************************
						'ACERVO - QUANTIDADE DE EXEMPLARES
						'*************************************
                        if xmlCampos.nodeName = "ACERVO" then
                            acervo_qtdes = ""
                            acervo_desc = xmlCampos.attributes.getNamedItem("Descricao").value 
                            ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"&acervo_desc&"</td>"
                            ficha = ficha & "<td class='esquerda td_ficha_dir'>"
                            For Each xmlSubCampos In xmlCampos.childNodes
                                if acervo_qtdes <> "" then
                                    ficha = ficha & "<br/><br/>"    
                                end if
                                acervo_qtdes = xmlSubCampos.attributes.getNamedItem("Descricao").value
                                ficha = ficha & acervo_qtdes
                            Next	
                            ficha = ficha & "</td></tr>"    
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
							if titulo_autor_sem_capa <> "" then
								titulo_autor_sem_capa = Replace(titulo_autor_sem_capa, "#COD#", codigo_atual)
							end if
							
							if global_exibe_capa = 1 then
								if CStr(xmlCampos.attributes.getNamedItem("Capa").value) = "1" then
									Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
									Response.Write "<table style='display:inline-table'><tr><td>"
									Response.Write "<br />"
									if url_imagem_referencia <> "" then
										Response.Write "<a href='" & url_imagem_referencia & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>"
									end if
									Response.Write "<img alt='' src='asp/capa.asp?obra="&codigo_atual&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'>"
									if url_imagem_referencia <> "" then
										Response.Write "</a>"
									end if
									Resposne.Write "<br />&nbsp;"
									Response.Write "</td></tr></table>"
									Response.Write "</td>"
								else
    								Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
									Response.Write "<table style='display: inline-table;'><tr><td>"
									Response.Write "<br/>"
									Response.Write  titulo_autor_sem_capa
									Response.Write "<div class='capa_fantasia icon_capa'></div>"
									Response.Write "<br/>"
									Response.Write "</td></tr></table>"
									Response.Write "</td>"
								end if
							end if
							Response.Write "<td class='td_center_top td_grid_ficha_background;' style='border-bottom: 1px solid #ccc;'>"
						end if
						'*************************************
						'MATERIAL
						'*************************************
						if xmlCampos.nodeName = "MIDIA" then
							material = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							material_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & material_desc
							ficha = ficha & "</td><td class='esquerda td_ficha_dir'>"&material&"</td></tr>"
						end if
						'*************************************
						'NORMA E NÚMERO
						'*************************************
						if xmlCampos.nodeName = "NORMA" then
							norma = RemoveTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							norma_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							norma_f = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							pos_dec = Registro - 1

							if(sOrigem = "Pesquisa")then
								total = GetSessao(global_cookie,"nrows"&iIndexSrv)
								norma = "<b><a class='link_classic2' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&","&pagina&",'resultado',"&Tipo&");"">"&norma_f &"</a></b>"
							elseif(sOrigem = "MySel")then
								total = iNumSel
								norma = "<b><a class='link_classic2' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'sels',"&Tipo&");"">"&norma_f &"</a></b>"
							elseif(sOrigem = "BibCurso")then
								total = Resultado
								norma = "<b><a class='link_classic2' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'bibcurso',"&Tipo&");"">"&norma_f &"</a></b>"
							elseif(sOrigem = "favoritos") then
								total = quantidade
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'favoritos',"&Tipo&");"">"&titulo_f &"</a>"
							end if
						
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & norma_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"
							ficha = ficha & norma
							ficha = ficha & "</td></tr>"	
						end if
						'*************************************
						'APELIDO
						'*************************************
						if xmlCampos.nodeName = "APELIDO" then
							apelido = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							apelido_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & apelido_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&apelido&"</td></tr>"
						end if
						'*************************************
						'DATA ASSINATURA
						'*************************************
						if xmlCampos.nodeName = "DATA_ASSINATURA" then
							data_assinatura = xmlCampos.attributes.getNamedItem("Valor").value
							data_assinatura_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & data_assinatura_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&data_assinatura&"</td></tr>"
						end if
						'*************************************
						'EMENTA
						'*************************************
						if xmlCampos.nodeName = "EMENTA" then
							ementa = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							ementa_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & ementa_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>" & replace(ementa,chr(10),"<br/>") & "</td></tr>"
						end if
						'*************************************
						'DATA DE PUPLICAÇÃO
						'*************************************
						if xmlCampos.nodeName = "DATA_PUBLICACAO" then
							data_publicacao = xmlCampos.attributes.getNamedItem("Valor").value
							data_publicacao_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & data_publicacao_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&data_publicacao&"</td></tr>"
						end if						
						'*************************************
						'ORGAO DE ORIGEM
						'*************************************
						if xmlCampos.nodeName = "ORGAO_ORIGEM" then
							desc_orgao = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							orgao_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							orgao_busca = RemoveTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							orgao_codigo = xmlCampos.attributes.getNamedItem("Codigo").value							
							orgao_formatado = replace(replace(replace(replace(orgao_busca," ","_"),"<",""),">",""),"'","\'")
							
							'TJ-RJ
							if (global_numero_serie = 4794) or (global_numero_serie = 5613) then
								hint = "Buscar todas legislações desta instituição"
							else
								hint = getTermo(global_idioma, 1559, "Buscar todos os registros desta instituição", 0)&"..."
							end if
							orgao = "<a class='link_classic2' title='"&hint&"...' href=""javascript:LinkBuscaOrgao(parent.hiddenFrame.modo_busca,"&orgao_codigo&",'"&orgao_formatado&"',"&iIndexSrv&");"">"&desc_orgao&"</a>"
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & orgao_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&orgao&"</td></tr>"
						end if
						'*************************************
						'ESFERA
						'*************************************
						if xmlCampos.nodeName = "ESFERA" then
							esfera = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							esfera_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & esfera_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&esfera&"</td></tr>"
						end if
						'*************************************
						'SITUAÇÃO DA LEGISLAÇÃO
						'*************************************
						if xmlCampos.nodeName = "SITUACAO_LEGISLACAO" then
							leg_situacao = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							leg_situacao_desc = xmlCampos.attributes.getNamedItem("Descricao").value

							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & leg_situacao_desc
							
							css_situacao =  "td_ficha_dir"
							
							if (CStr(xmlCampos.attributes.getNamedItem("Destaque").value) = "1") then
								css_situacao = css_situacao&" td_leg_valor_detalhe"
							end if
							
							ficha = ficha & "</td><td class='"&css_situacao&" esquerda'>"&leg_situacao&"</td></tr>"
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
							ficha = ficha & "<td class='td_ficha_esq direita'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'*************************************
								'SE EXIBE TODOS OS ASSUNTOS NA FICHA
								'*************************************
								if (exibe_todos_ass = "1") then
									tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
									valor_assunto = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
									busca_assunto = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
									desc_princ_assunto = replace(replace(replace(busca_assunto," ","_"),"&#39;","_#39;"),"'","\'")
									seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)

									nome_funcao = "LinkBuscaAssunto"
									'TJ-RJ
									if (global_numero_serie = 4794) or (global_numero_serie = 5613) then
										hint = "Buscar todas legislações deste assunto"
									else
										hint = getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)
									end if

									if (trim(hint) <> "") then
										hint = hint & "..."
									end if

									if (sOrigem <> "emprestimo") and (xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1") and (config_habilita_autoridades = 1) then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assInfo = "<span class='assuntoImg span_imagem icon_16' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</span>"
										else
											assInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
										end if
									else
										assInfo = ""
									end if
									temp = seq_assunto&" <a class='link_classic2' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_assunto&",'"&desc_princ_assunto&"',"&iIndexSrv&");"">"&replace(valor_assunto,chr(10),"<br />") &"</a>"
                                    identificadoresInfo = ""
                                    For Each xmlIdentificadores in xmlSubCampos.childNodes                                
	                                    %><!--#include file ="identificadoresDetalhe.asp"--><%
                                    next
                                    assInfo = assInfo & identificadoresInfo
									if assunto <> "" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & "&nbsp;" & temp & "&nbsp;" & assInfo
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
										end if
									else
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & temp & "&nbsp;" & assInfo
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
										end if
									end if
									seq = seq + 1
								'*************************************
								'SE EXIBE SOMENTE PRIMEIRO ASSUNTO
								'*************************************
								else
									desc_assunto = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
                                    tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									assunto_desc = xmlSubCampos.attributes.getNamedItem("Descricao").value
									assunto_busca = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
									assunto_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							
									assunto_formatado = replace(replace(replace(replace(assunto_busca," ","_"),"<",""),">",""),"'","\'")
									if (sOrigem <> "emprestimo") and (xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1") and (config_habilita_autoridades = 1) then
										assuntoInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&sequencial&"' onClick=LinkAutInfo('linkAutInfo',"&assunto_codigo&",'"&assunto_formatado&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
									else
										assuntoInfo = ""
									end if
									identificadoresInfo = ""
                                    For Each xmlIdentificadores in xmlSubCampos.childNodes                                
	                                    %><!--#include file ="identificadoresDetalhe.asp"--><%
                                    next
                                    assuntoInfo = assuntoInfo & identificadoresInfo
									'TJ-RJ
									if (global_numero_serie = 4794) or (global_numero_serie = 5613) then
										hint = "Buscar todas legislações deste assunto"
									else
										hint = getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)
									end if
									assunto = "<table class='autLink remover_bordas_padding'><tr><td id='assLinkTab"&sequencial&"' class='td_left_middle'><a class='link_custom' title='"&hint&"...' href=""javascript:LinkBuscaAssunto(parent.hiddenFrame.modo_busca,"&assunto_codigo&",'"&assunto_formatado&"',"&iIndexSrv&");"">"&desc_assunto&"</a>&nbsp;</td>"&assuntoInfo&"</tr></table>"
									ficha = ficha & "<td class='td_ficha_dir esquerda'>"&assunto&"</td></tr>"
									if len(trim(desc_assunto)) > 85 then
										ficha = ficha &"<script>redimensiona('assLinkTab"&sequencial&"',476);</script>"
									end if
								end if
							Next
							if (exibe_todos_ass = "1") then
								ficha = ficha & "<td class='td_ficha_dir esquerda'>" & assunto & "</td>"
								ficha = ficha & "</tr>"
							end if						
						end if
						'*************************************
						'SITES
						'*************************************
						if xmlCampos.nodeName = "SITE" then
							site_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							site_c = ""
							For Each xmlSubCampos In xmlCampos.childNodes
								codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
								plataforma = xmlSubCampos.attributes.getNamedItem("Plataforma").value
								site = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								site_url = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
                                dspace = xmlSubCampos.attributes.getNamedItem("DSpace").value
								if site_c <> "" then
									site_c = site_c & "<br/>"
								end if
								if (plataforma <> "") then
									site_c = site_c & "<a class='link_classic2' href='javascript:abreMidiaEspecifica(" & Tipo & "," & codigo_atual & "," & codMidia & "," & iIndexSrv & ")'>" & site & "</a>"
								else
									site_c = site_c & "<a class='link_classic2' href='" & site_url & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>" & site & "</a>"
								end if
							Next	
														
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & site_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda' style='word-break: break-all;'>" & site_c & "</td></tr>"
						end if
					'*************************************
					'FICHA DE ANALÍTICA
					'*************************************						
					elseif cStr(Tipo) = "3" then
						'*************************************
						'CODIGO
						'*************************************
						if xmlCampos.nodeName = "CODIGO" then
							codigo_atual = xmlCampos.attributes.getNamedItem("Valor").value
							if titulo_autor_sem_capa <> "" then
								titulo_autor_sem_capa = Replace(titulo_autor_sem_capa, "#COD#", codigo_atual)
							end if
							
							if global_exibe_capa = 1 then
								if CStr(xmlCampos.attributes.getNamedItem("Capa").value) = "1" then
									Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
									Response.Write "<table style='display: inline-table'><tr><td>"
									Response.Write "<br />"
									if url_imagem_referencia <> "" then
										Response.Write "<a href='" & url_imagem_referencia & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>"
									end if
									Response.Write "<img alt='' src='asp/capa.asp?obra="&codigo_atual&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'>"
									if url_imagem_referencia <> "" then
										Response.Write "</a>"
									end if
									Response.Write "<br />&nbsp;"
									Response.Write "</td></tr></table>"
									Response.Write "</td>"
								else
									if (codigo_material = "46" or codigo_material = "35") then
										Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
										Response.Write "<table style='display: inline-table'><tr><td>"
										Response.Write "<br />"
										if url_imagem_referencia <> "" then
											Response.Write "<a href='" & url_imagem_referencia & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia&",1,"&dspace&");'>"
										end if
										Response.Write "<div class='capa_audio icon_capa'></div>" 
										if url_imagem_referencia <> "" then
											Response.Write "</a>"
										end if
										Response.Write "<br />&nbsp;"
										Response.Write "</td></tr></table>"
										Response.Write "</td>"
									else
										Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
										Response.Write "<table style='display: inline-table'><tr><td>"
										Response.Write "<br/>"
										Response.Write  titulo_autor_sem_capa
										Response.Write "<div class='capa_fantasia icon_capa'></div>"
										Response.Write "<br/>"
										Response.Write "</td></tr></table>"
										Response.Write "</td>"
									end if
								end if
							end if
							Response.Write "<td class='td_center_top td_grid_ficha_background' style='border-bottom: 1px solid #ccc;'>"
						end if
						'*************************************
						'MATERIAL
						'*************************************
						if xmlCampos.nodeName = "MIDIA" then
							material = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							material_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & material_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&material&"</td></tr>"
						end if
						'*************************************
						'NUMERO DE CHAMADA
						'*************************************
						if xmlCampos.nodeName = "CHAMADA" then
							chamada = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							chamada_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & chamada_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&chamada&"</td></tr>"
						end if
						'*************************************
						'AUTOR (ENTRADA PRINCIPAL)
						'*************************************
						if xmlCampos.nodeName = "ENT_PRINC" then
							desc_autor = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							autor_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							autor_busca = RemoveTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							autor_tipo = xmlCampos.attributes.getNamedItem("Tipo").value
							autor_codigo = xmlCampos.attributes.getNamedItem("Codigo").value							
							nome_funcao = "LinkBuscaAutor"
							autor_formatado = replace(replace(replace(replace(autor_busca," ","_"),"<",""),">",""),"'","\'")
							if (sOrigem <> "emprestimo") and (xmlCampos.attributes.getNamedItem("Link_Autoridades").value = "1") and (config_habilita_autoridades = 1) then
								autInfo = "<td class='autImg'><div class='autImg' id='autImg"&sequencial&"' onClick=LinkAutInfo('linkAutInfo',"&autor_codigo&",'"&autor_formatado&"',"&autor_tipo&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1571, "Mostrar informações sobre este autor", 0)&"...'>&nbsp;</div></td>"
							else
								autInfo = ""
							end if
                            identificadoresInfo = ""
                            For Each xmlIdentificadores in xmlCampos.childNodes                                
	                            %><!--#include file ="identificadoresDetalhe.asp"--><%
                            next
                            autInfo = autInfo & identificadoresInfo
							autor = "<table class='autLink remover_bordas_padding'><tr><td id='autLinkTab"&sequencial&"' class='td_left_middle'><a class='link_custom' title='"&getTermo(global_idioma, 1558, "Buscar todos os registros deste autor", 0)&"...' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&autor_codigo&",'"&autor_formatado&"',"&iIndexSrv&");"">"&desc_autor&"</a>&nbsp;</td>"&autInfo&"</tr></table>"
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & autor_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&autor&"</td></tr>"
							if len(trim(desc_autor)) > 85 then
								ficha = ficha &"<script>redimensiona('autLinkTab"&sequencial&"',476);</script>"
							end if							
						end if
                        '*************************************
						'LOCALIZAÇÃO (BN)
						'*************************************
						if xmlCampos.nodeName = "LOCALIZACAO" then
							loc = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							loc_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & loc_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&loc&"</td></tr>"
						end if
						'*************************************
						'TITULO
						'*************************************
						if xmlCampos.nodeName = "TITULO" then
							titulo = xmlCampos.attributes.getNamedItem("Titulo").value
							titulo_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							complemento = " " & TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Complemento").value)
							titulo_f = Replace(Replace(titulo,"<","&lt;"),">","&gt;")							
							pos_dec = Registro - 1
							
							if(sOrigem = "Pesquisa") then
								total = GetSessao(global_cookie,"nrows"&iIndexSrv)
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&","&pagina&",'resultado',"&Tipo&");"">"&titulo_f &"</a>"
							elseif(sOrigem = "MySel") then
								total = iNumSel
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'sels',"&Tipo&");"">"&titulo_f &"</a>"
							elseif(sOrigem = "BibCurso") then
								total = Resultado
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'bibcurso',"&Tipo&");"">"&titulo_f &"</a>"
							elseif(sOrigem = "favoritos") then
								total = quantidade
								titulo = "<a class='link_custom_negrito' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'favoritos',"&Tipo&");"">"&titulo_f &"</a>"
							end if
									
							titulo = TrocaTagMarcador(titulo)
									
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & titulo_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"
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
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & data_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&data&"</td></tr>"
						end if
						'************************************************************
						' Nº ÁLBUM
						'Adequeação: apenas IMS
						'************************************************************
						if xmlCampos.nodeName = "COMPLEMENTO1" and (global_numero_serie = 4516) then
							complemento = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
							desc_complemento = xmlCampos.attributes.getNamedItem("Descricao").value
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & desc_complemento
							ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&complemento&"</td></tr>"
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
							ficha = ficha & "<td class='td_ficha_esq direita'>"&desc_assunto&"</td>"
							For Each xmlSubCampos In xmlCampos.childNodes
								'*************************************
								'SE EXIBE TODOS OS ASSUNTOS NA FICHA
								'*************************************
								if (exibe_todos_ass = "1") then
									tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									codigo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Codigo").value)
									valor_assunto = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
									busca_assunto = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
									desc_princ_assunto = replace(replace(replace(busca_assunto," ","_"),"&#39;","_#39;"),"'","\'")
									seq_assunto = trim(xmlSubCampos.attributes.getNamedItem("Seq").value)
									nome_funcao = "LinkBuscaAssunto"
									hint = getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"..."
									if (sOrigem <> "emprestimo") and (xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1") and (config_habilita_autoridades = 1) then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assInfo = "<span class='assuntoImg span_imagem icon_16' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</span>"
										else
											assInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&seq&"' onClick=LinkAutInfo('detalhes',"&codigo_assunto&",'"&desc_princ_assunto&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
										end if
									else
										assInfo = ""
									end if
									temp = seq_assunto&" <a class='link_classic2' title='"&hint&"' href=""javascript:"&nome_funcao&"(parent.hiddenFrame.modo_busca,"&codigo_assunto&",'"&desc_princ_assunto&"',"&iIndexSrv&");"">"&replace(valor_assunto,chr(10),"<br />") &"</a>"
                                    identificadoresInfo = ""
                                    For Each xmlIdentificadores in xmlSubCampos.childNodes                                
	                                    %><!--#include file ="identificadoresDetalhe.asp"--><%
                                    next
                                    assInfo = assInfo & identificadoresInfo
									if assunto <> "" then
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & "&nbsp;" & temp & "&nbsp;" & assInfo
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px; padding: 0'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
										end if
									else
										'Adequação ABL - Exibir itens um em frente do outro
										if ((global_numero_serie = 2372) or (global_numero_serie = 2635)) then
											assunto = assunto & temp & "&nbsp;" & assInfo
										else
											assunto = assunto & "<table class='autLink' style='border-spacing: 1px ; padding: 0'><tr><td id='assLinkTab"&seq&"' class='td_left_middle'>" & temp & "&nbsp;</td>" & assInfo & "</tr></table>"
										end if
									end if
									seq = seq + 1
								'*************************************
								'SE EXIBE SOMENTE PRIMEIRO ASSUNTO
								'*************************************
								else
									desc_assunto = TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
                                    tipo_assunto = trim(xmlSubCampos.attributes.getNamedItem("Tipo").value)
									assunto_desc = xmlSubCampos.attributes.getNamedItem("Descricao").value
									assunto_busca = RemoveTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)
									assunto_codigo = xmlSubCampos.attributes.getNamedItem("Codigo").value							
									assunto_formatado = replace(replace(replace(replace(assunto_busca," ","_"),"<",""),">",""),"'","\'")
									if (sOrigem <> "emprestimo") and (xmlSubCampos.attributes.getNamedItem("Link_Autoridades").value = "1") and (config_habilita_autoridades = 1) then
										assuntoInfo = "<td class='assuntoImg'><div class='assuntoImg' id='assuntoImg"&sequencial&"' onClick=LinkAutInfo('linkAutInfo',"&assunto_codigo&",'"&assunto_formatado&"',"&tipo_assunto&") style='cursor:pointer;' title='"&getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"...'>&nbsp;</div></td>"
									else
										assuntoInfo = ""
									end if
                                    identificadoresInfo = ""
                                    For Each xmlIdentificadores in xmlSubCampos.childNodes                                
	                                    %><!--#include file ="identificadoresDetalhe.asp"--><%
                                    next
                                    assuntoInfo = assuntoInfo & identificadoresInfo
									assunto = "<table class='autLink remover_bordas_padding'><tr><td id='assLinkTab"&sequencial&"' class='td_left_middle'><a class='link_custom' title='"&getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"...' href=""javascript:LinkBuscaAssunto(parent.hiddenFrame.modo_busca,"&assunto_codigo&",'"&assunto_formatado&"',"&iIndexSrv&");"">"&desc_assunto&"</a>&nbsp;</td>"&assuntoInfo&"</tr></table>"
									ficha = ficha & "<td class='td_ficha_dir esquerda'>"&assunto&"</td></tr>"
									if len(trim(desc_assunto)) > 85 then
										ficha = ficha &"<script>redimensiona('assLinkTab"&sequencial&"',476);</script>"
									end if
								end if
							Next
							if (exibe_todos_ass = "1") then
								ficha = ficha & "<td class='td_ficha_dir esquerda'>" & assunto & "</td>"
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
										sDesc_fonte    = xmlFonte.attributes.getNamedItem("DESCRICAO").value
										sTit_fonte     = TrocaTagMarcador(xmlFonte.attributes.getNamedItem("TITULO").value)
										sComp_fonte    = TrocaTagMarcador(xmlFonte.attributes.getNamedItem("COMPLEMENTO").value)
										sPreComp_fonte = TrocaTagMarcador(xmlFonte.attributes.getNamedItem("PRE_COMP").value)
										sTipo_fonte    = xmlFonte.attributes.getNamedItem("TIPO").value
										sCod_fonte     = xmlFonte.attributes.getNamedItem("CODIGO").value
																		
										if cStr(sTipo_Fonte) = "0" then
											sCodEx	   = xmlFonte.attributes.getNamedItem("EXEMPLAR").value
											sTit_fonte = "<a class='link_classic2' title='"&getTermo(global_idioma, 1568, "Ver detalhes do periódico", 0)&"...' href=""javascript:LinkDetalhesPeriodico(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&sCod_fonte&",1,'periodico',"&sCodEx&",0);"">"&sTit_fonte&"</a>"
										else
											if cStr(sTipo_Fonte) = "20" then
												sTit_fonte = "<a class='link_classic2' title='"&getTermo(global_idioma, 1568, "Ver detalhes do periódico", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&sCod_fonte&","&pagina&",'periodico',"&sTipo_Fonte&");"">"&sTit_fonte&"</a>"
											else
												sTit_fonte = "<a class='link_classic2' title='"&getTermo(global_idioma, 1567, "Ver detalhes da obra", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&sCod_fonte&","&pagina&",'periodico',"&sTipo_Fonte&");"">"&sTit_fonte&"</a>"
											end if
										end if
																							
										ficha = ficha & "<tr>"
										ficha = ficha & "<td class='td_ficha_esq direita'>"
										ficha = ficha & sDesc_fonte
										ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"
										ficha = ficha & sPreComp_fonte&sTit_fonte&sComp_fonte
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
							For Each xmlSubCampos In xmlCampos.childNodes
								codMidia = xmlSubCampos.attributes.getNamedItem("Codigo").value
								plataforma = xmlSubCampos.attributes.getNamedItem("Plataforma").value
								site = TrocaTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
								site_url = RemoveTagMarcador(trim(xmlSubCampos.attributes.getNamedItem("Valor").value))
                                dspace = xmlSubCampos.attributes.getNamedItem("DSpace").value
								if site_c <> "" then
									site_c = site_c & "<br/>"
								end if
								if (plataforma <> "") then
									site_c = site_c & "<a class='link_classic2' href='javascript:abreMidiaEspecifica(" & Tipo & "," & codigo_atual & "," & codMidia & "," & iIndexSrv & ")'>" & site & "</a>"
								else
									site_c = site_c & "<a class='link_classic2' href='" & site_url & "' target='_blank' onclick='ContarAcesso("&codigo_atual&","&codMidia& ",1,"&dspace&");'>" & site & "</a>"
								end if
							Next	
														
							ficha = ficha & "<tr>"
							ficha = ficha & "<td class='td_ficha_esq direita'>"
							ficha = ficha & site_desc
							ficha = ficha & "</td><td class='td_ficha_dir esquerda' style='word-break: break-all;'>" & site_c & "</td></tr>"
						end if

					end if

					'*************************************
					'AVALIAÇÃO (RANKING)
					'*************************************
					if (sOrigem <> "emprestimo") then	
						if xmlCampos.nodeName = "avaliacao" then
							'Verifica se avaliação está habilitada
							if (global_avaliacao_online = 1) then
								fichaAvaliacao = ""
								mediaAvaliacao = xmlCampos.attributes.getNamedItem("media").value
								quantidadeAvaliacao = xmlCampos.attributes.getNamedItem("quantidade").value
								fichaAvaliacao = "<p style='display: inline; height: 15px;' class='td_left_middle'>" & ImprimeEstrelas(mediaAvaliacao,false) & "</p>"
							
								if quantidadeAvaliacao = 0 then
									if global_avaliacao_autenticada = 1 then
										if Session("Logado")= "sim" then
											linkAvaliacao = "<a class='link_classic2' title='"& getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"&CodigoUsuario=" & Session("codigo_usuario") & "','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
											fichaAvaliacao = fichaAvaliacao & linkAvaliacao & "<p style='display: inline; margin-left: 10px; height: 15px;' class='td_left_middle'>" & getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "</p></a>"
										else
											linkAvaliacao = "<a class='link_classic2' title='"& getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "' style='cursor:pointer;' onclick=""abreLogin('avaliacao_votar','"&getTermo(global_idioma, 6689, "Avaliação on-line", 0)&"','&codigo_obra="&codigo_atual&"',true,true);"">"
											fichaAvaliacao = fichaAvaliacao & linkAvaliacao & "<p style='display: inline; margin-left: 10px; height: 15px;' class='td_left_middle'>" & getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "</p></a>"
										end if
									else
										linkAvaliacao = "<a class='link_classic2' title='"& getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"&CodigoUsuario=0','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
										fichaAvaliacao = fichaAvaliacao & linkAvaliacao & "<p style='display: inline; margin-left: 10px; height: 15px;' class='td_left_middle'>" & getTermoHtml(global_idioma, 6700, "Seja o primeiro a avaliar ", 0) & "</p></a>"
									end if
								elseif quantidadeAvaliacao = 1 then
									if global_avaliacao_autenticada = 1 then
										if Session("Logado")= "sim" then
											linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"&CodigoUsuario=" & Session("codigo_usuario") & "','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
											linkMediaAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) &"' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,false,true);"">"
											fichaAvaliacao = fichaAvaliacao & "<p style='display: inline; margin-left: 10px; height: 15px;' class='td_left_middle'>" & linkMediaAvaliacao & getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</p></a>"
										else 
											linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;' onclick=""abreLogin('avaliacao_votar','"&getTermo(global_idioma, 6689, "Avaliação on-line", 0)&"','&codigo_obra="&codigo_atual&"',true,true);"">"
											linkMediaAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) &"' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,false,true);"">"
											fichaAvaliacao = fichaAvaliacao & "<p style='display: inline; margin-left: 10px; height: 15px;' class='td_left_middle'>" & linkMediaAvaliacao & getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</p></a>"
										end if
									else
										linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"&CodigoUsuario=0','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
										linkMediaAvaliacao = "<a class='link_classic2' title='" & getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) &"' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
										fichaAvaliacao = fichaAvaliacao & "<p style='display: inline; margin-left: 10px; height: 15px;' class='td_left_middle'>" & linkMediaAvaliacao & getTermo(global_idioma, 6690, "Uma pessoa avaliou", 0) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</p></a>"
									end if
								else
									if global_avaliacao_autenticada = 1 then
										if Session("Logado")= "sim" then
											linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"&CodigoUsuario=" & Session("codigo_usuario") & "','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
											linkMediaAvaliacao = "<a class='link_classic2' title='"& Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) &"' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
											fichaAvaliacao = fichaAvaliacao & "<p style='display: inline; margin-left: 10px; height: 15px;' class='td_left_middle'>" & linkMediaAvaliacao & Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</p></a>"
										else
											linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;' onclick=""abreLogin('avaliacao_votar','"&getTermo(global_idioma, 6689, "Avaliação on-line", 0)&"','&codigo_obra="&codigo_atual&"',true,true);"">"
											linkMediaAvaliacao = "<a class='link_classic2' title='"& Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao)&"' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
											fichaAvaliacao = fichaAvaliacao & "<p style='display: inline; margin-left: 10px; height: 15px;' class='td_left_middle'>" & linkMediaAvaliacao & Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</p></a>"
										end if
									else

										linkAvaliacao = "<a class='link_classic2' title='"& getTermo(global_idioma, 6692, "Avaliar", 0) & "' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_votar.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"&CodigoUsuario=0','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',300,300,true,true);"">"
										linkMediaAvaliacao = "<a class='link_classic2' title='"& Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao)&"' style='cursor:pointer;' onclick=""abrePopup('asp/avaliacao_exibir_resultados.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&Codigo="&codigo_atual&"','" & getTermo(global_idioma, 6689, "Avaliação on-line", 0) & "',400,300,true,true);"">"
										fichaAvaliacao = fichaAvaliacao & "<p style='display: inline; margin-left: 10px; height: 15px;' class='td_left_middle'>" & linkMediaAvaliacao & Format(getTermo(global_idioma, 6691, "%s pessoas avaliaram", 0),quantidadeAvaliacao) & "</a>" & " - " & linkAvaliacao & getTermo(global_idioma, 6692, "Avaliar", 0) & "</p></a>"
									end if
								end if
							end if
						end if
					
						'*************************************
						'SERVIÇOS HABILITADOS PARA O ITEM
						'*************************************
						if (global_esconde_menu = 0) then
							if xmlCampos.nodeName = "LINK_SELECIONAR" then
								serv_selecao = true
							end if
							if xmlCampos.nodeName = "LINK_RESERVAR"  then
								serv_reserva = (repositorio_institucional = 0)
							end if
							if xmlCampos.nodeName = "LINK_AQUISICAO" then 
								serv_aquisicao = (repositorio_institucional = 0)
							end if
						end if

						if xmlCampos.nodeName = "LINK_MIDIAS" then
							serv_midias = true
							serv_midias_desc = xmlCampos.attributes.getNamedItem("Descricao").value
							serv_midias_mus = xmlCampos.attributes.getNamedItem("Audio").value
						end if
						if xmlCampos.nodeName = "LINK_EXEMPLAR" then
							serv_exemplar = true
						end if
						if xmlCampos.nodeName = "LINK_ANALITICA" then
							serv_analitica = true
							serv_analitica_desc = xmlCampos.attributes.getNamedItem("Descricao").value
						end if
						if xmlCampos.nodeName = "LINK_REF_BIB" then
							serv_referencia = true
							serv_referencia_desc = xmlCampos.attributes.getNamedItem("Descricao").value
						end if
					end if
				Next

				'*************************************
				'REDES SOCIAIS
				'*************************************
						
				divRedesSociais = "<div class='td_left_middle div-redes-sociais' style='height: 18px; width:100px;color: #00395B; float:left;padding-left:10px;'>"
							
				if (sOrigem <> "emprestimo") then
					linkRedesSociais = Session("baseUrl") & "info.asp?c="&codigo_atual
					if global_facebook_curtir = 1 then
						divRedesSociais = "<div class='td_left_middle div-redes-sociais' style='height: 18px; width:200px;color: #00395B; float:left;padding-left:10px;'>"
						' Utilização do Curtir do Facebook via API
						if global_facebook_iframe <> 1 then	
							divRedesSociais = divRedesSociais & "<div class='div-facebook' style='height: 18px; width:110px;float:left;visibility:hidden;overflow:hidden;'>"
							divRedesSociais = divRedesSociais & "<fb:like href="&linkRedesSociais&" send='false' show_faces='false' layout='button_count' colorscheme='dark'></fb:like>&nbsp;&nbsp;"
							divRedesSociais = divRedesSociais & "</div>"
						' Utilização do Curtir do Facebook via iFRAME
						else
							divRedesSociais = divRedesSociais & "<div class='div-facebook' style='height: 18px; width:100px;float:left;'>"
							divRedesSociais = divRedesSociais & "<iframe src=""//www.facebook.com/plugins/like.php?href=" & Server.URLEncode(linkRedesSociais) & "&amp;" & _
								"layout=button_count&amp;width=100&amp;show_faces=false&amp;font&amp;colorscheme=light&amp;action=like&amp;height=20"" scrolling=""no"" " & _
								"frameborder=""0"" style=""border:none; overflow:hidden; width:100px; height:20px;"" allowTransparency=""true""></iframe>&nbsp;&nbsp;"
							divRedesSociais = divRedesSociais & "</div>"

						end if

					end if
					
					if global_twitter = 1 then
						divRedesSociais = divRedesSociais & "<div class='div-twitter' style='height: 18px; width:90px;float:right;'>"
						divRedesSociais = divRedesSociais & "<a href='https://twitter.com/share' class='twitter-share-button' data-counturl='"&linkRedesSociais&"' data-url='"&linkRedesSociais&"' data-text='"&getTermoHtml(global_idioma, 7211, "Confira", 0)&": '>Tweet</a>"
						divRedesSociais = divRedesSociais & "</div>"
					end if
				end if	
				divRedesSociais = divRedesSociais & "</div>"

				areaAvaliacaoRedesSociais = "<tr><td colspan='2' class='esquerda' style='height: 30px;'>" & _
					"<div class='td_left_middle' style='height: 15px; color: #00395B; float:left;'>" & _
					fichaAvaliacao & "</div>" & divRedesSociais & "</td></tr>"

				ficha = ficha & areaAvaliacaoRedesSociais

				if tem_chamada = false then
					ficha = Replace(ficha,"colspan=2","")
				end if
				Response.Write "<br /><table style='border-spacing: 2px; padding: 1px; width: 98%; display: inline-table;' class='td_grid_ficha_background'>"&ficha&"</table><br />"
				Response.Write "</td>"
				
				'*************************************
				' CELULA COM SERVIÇOS
				'*************************************
				if (sOrigem <> "emprestimo") then
					if(sOrigem = "Pesquisa") then
						total = GetSessao(global_cookie,"nrows"&iIndexSrv)
					elseif(sOrigem = "MySel") then
						total = iNumSel
					elseif(sOrigem = "BibCurso") then
						total = Resultado
					elseif(sOrigem = "favoritos") then
						total = quantidade
					end if

					pos_dec = Registro - 1
					Response.Write "<td class='td_center_top_padding td_grid_ficha_background td_grid_ficha_borda' style='width: 140px; border-bottom: 1px solid #ccc;'>"
					Response.Write "<table style='border-spacing: 1px; padding: 0px; background-color: #ffffff; display: inline-block;'><tr>"
					'*************************************
					'SELECIONAR
					'*************************************
					if (serv_selecao = true) and ((sOrigem = "Pesquisa") or (sOrigem = "BibCurso")) then
						s_sel = "<input id='srv"&iIndexSrv&"_cksel"&codigo_atual&"' type='checkbox' style='cursor: pointer; margin-right: 10px' value='"&codigo_atual&"."&Tipo&"'>"
						if (global_numero_serie = 5631) and (codigo_material = "33") then
							s_sel = s_sel & "<a class='link_serv' title='"&getTermo(global_idioma, 9575, "Selecionar imagem", 0)&"' style='cursor:pointer; display: inline-block;'  onclick=""ficha_selecionar("&iIndexSrv&","&codigo_atual&");"">"
							s_sel = s_sel & "&nbsp;<div style='float:right; width:100px;'>"&getTermo(global_idioma, 9575, "Selecionar imagem", 0)&"</div></a>"
						else
							s_sel = s_sel & "<a class='link_serv' title='"&getTermo(global_idioma, 1199, "Selecionar item", 0)&"' style='cursor:pointer;' onclick=""ficha_selecionar("&iIndexSrv&","&codigo_atual&");"">"
							s_sel = s_sel & "&nbsp;"&getTermo(global_idioma, 128, "Selecionar", 0)&"</a>"
						end if
						Response.Write "<td class='esquerda td_ficha_serv' id='srv"&iIndexSrv&"_s_sel"&codigo_atual&"' onMouseOver=servicosResultOver('srv"&iIndexSrv&"_s_sel"&codigo_atual&"'); onMouseOut=servicosResultOut('srv"&iIndexSrv&"_s_sel"&codigo_atual&"');>"&s_sel&"</td></tr>"
					end if
					'*************************************
					'DETALHES
					'*************************************                                                                                                                                                                                                                           
					if(sOrigem = "Pesquisa") then
						s_det ="<a class='link_serv' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' style='cursor:pointer;' onClick=LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&","&pagina&",'resultado',"&Tipo&") ><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-search' data-icon='search'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 1032, "Detalhes", 0)&"</a>"
					elseif(sOrigem = "MySel") then
						s_det ="<a class='link_serv' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' style='cursor:pointer;' onClick=LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'sels',"&Tipo&") ><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-search' data-icon='search'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 1032, "Detalhes", 0)&"</a>"
					elseif(sOrigem = "BibCurso") then
						s_det ="<a class='link_serv' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' style='cursor:pointer;' onClick=LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'bibcurso',"&Tipo&") ><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-search' data-icon='search'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 1032, "Detalhes", 0)&"</a>"
					elseif(sOrigem = "favoritos") then
						s_det ="<a class='link_serv' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' style='cursor:pointer;' onClick=LinkDetalhes(parent.hiddenFrame.modo_busca,"&total&","&pos_dec&","&codigo_atual&",1,'favoritos',"&Tipo&") ><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-search' data-icon='search'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 1032, "Detalhes", 0)&"</a>"
					end if
					Response.Write "<td class='esquerda td_ficha_serv' id='s_det"&codigo_atual&"' onMouseOver=servicosResultOver('s_det"&codigo_atual&"'); onMouseOut=servicosResultOut('s_det"&codigo_atual&"');>&nbsp;"&s_det&"</td></tr>"
					'*************************************
					'MIDIAS
					'*************************************
					if (serv_midias = true) then
						if global_numero_serie = 2635 then 'ABL_BRG
							s_midia = "<a class='link_serv' title='Visualizar doc. digital' style='cursor:pointer;' onclick=""abrePopup('asp/midia.asp?tipo="&Tipo&"&codigo="&codigo_atual&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','Doc. digital',500,490,true,true);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-digital' data-icon='digital'></span>&nbsp;&nbsp;Doc. digital&nbsp;</a>"
						else
							if serv_midias_mus = "1" then
								s_midia = "<a class='link_serv' title='"&getTermo(global_idioma, 1622, "Ouvir áudio", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/midia.asp?tipo="&Tipo&"&codigo="&codigo_atual&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&serv_midias_desc&"',500,490,true,true);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-audio' data-icon='audio'></span>&nbsp;&nbsp;"&serv_midias_desc&"&nbsp;</a>"
							else
								s_midia = "<a class='link_serv' title='"&getTermo(global_idioma, 1622, "Visualizar conteúdo digital", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/midia.asp?tipo="&Tipo&"&codigo="&codigo_atual&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&serv_midias_desc&"',500,490,true,true);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-digital' data-icon='digital'></span>&nbsp;&nbsp;"&serv_midias_desc&"&nbsp;</a>"
							end if
						end if
						Response.Write "<tr><td class=' esquerda td_ficha_serv' id='s_midia"&codigo_atual&"' onMouseOver=servicosResultOver('s_midia"&codigo_atual&"'); onMouseOut=servicosResultOut('s_midia"&codigo_atual&"');>&nbsp;"&s_midia&"</td>"
					end if
					'*************************************
					'ANALITICAS
					'*************************************
					if (serv_analitica = true) and (global_numero_serie <> 5516) and ((sOrigem = "Pesquisa") or (sOrigem = "BibCurso")) then
						s_ana = "<a class='link_serv' title='"&getTermo(global_idioma, 1402, "Exibir analíticas", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/artigos.asp?obra="&codigo_atual&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"&servidor="&iIndexSrv&"','"&serv_analitica_desc&"', 965, 450, false, true);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-document-b' data-icon='document-b'></span>&nbsp;&nbsp;"&serv_analitica_desc&"</a>"
						Response.Write "<tr><td class='esquerda td_ficha_serv' id='s_ana"&codigo_atual&"' onMouseOver=servicosResultOver('s_ana"&codigo_atual&"'); onMouseOut=servicosResultOut('s_ana"&codigo_atual&"');>&nbsp;"&s_ana&"</td>"
					end if
					'*************************************
					'EXEMPLARES
					'*************************************
					if (serv_exemplar = true) and ((sOrigem = "Pesquisa") or (sOrigem = "BibCurso")) then
						s_einfo = "<a class='link_serv' title='"&getTermo(global_idioma, 2670, "Exibir informações sobre exemplares", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/pop_exemplares.asp?c="&codigo_atual&"&t="&Tipo&"&qtde="&total&"&pagina="&pagina&"&posicao_vetor="&Registro&"&iIndexSrv="&iIndexSrv&"&biblioteca='+parent.hiddenFrame.geral_bib+'&projeto='+parent.hiddenFrame.iBusca_Projeto+'&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 337, "Exemplares", 0)&"',990,354,true,true);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-collection' data-icon='collection'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 337, "Exemplares", 0)&"</a>"
						Response.Write "<tr><td class='esquerda td_ficha_serv' id='s_einfo"&codigo_atual&"' onMouseOver=servicosResultOver('s_einfo"&codigo_atual&"'); onMouseOut=servicosResultOut('s_einfo"&codigo_atual&"');>&nbsp;"&s_einfo&"</td>"
					end if
					'*************************************
					'RESERVA
					'*************************************
					if serv_reserva = true and (Session("usuario_externo") = false) then
						if Session("Logado") = "sim" then
							if config_multi_servbib = 1 then
								if CStr(Session("Servidor_Logado")) = CStr(iIndexSrv) then
									s_res = "<a class='link_serv' title='"&getTermo(global_idioma, 1459, "Reservar", 0)&"' style='cursor:pointer;' onclick=""if(parent.hiddenFrame.modo_busca='telacheia'){parent.hiddenFrame.modo_busca='rapida';}abrePopup('asp/reserva.asp?veio_de=busca_principal&codigo_obra="&codigo_atual&"&tipo_obra="&Tipo&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 348, "Reserva", 0)&"',380,400,false,false);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-flag-b' data-icon='flag-b'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 1459, "Reservar", 0)&"</a>"
								else
									s_res = "<a class='link_serv' title='"&getTermo(global_idioma, 1407, "Efetue o login para reservar", 0)&"' style='cursor:pointer;' onclick=""abreLogin('reserva','"&getTermo(global_idioma, 348, "Reserva", 0)&"','&codigo_obra="&codigo_atual&"&tipo_obra="&Tipo&"&servidor="&iIndexSrv&"',false,true);""><span class='span_imagem div_imagem_right_3 icon_16 icon-small-flag-x'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 1459, "Reservar", 0)&"</a>"
								end if
							else
								s_res = "<a class='link_serv' title='"&getTermo(global_idioma, 1459, "Reservar", 0)&"' style='cursor:pointer;' onclick=""if(parent.hiddenFrame.modo_busca='telacheia'){parent.hiddenFrame.modo_busca='rapida';}abrePopup('asp/reserva.asp?veio_de=busca_principal&codigo_obra="&codigo_atual&"&tipo_obra="&Tipo&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 348, "Reserva", 0)&"',380,400,false,false);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-flag-b' data-icon='flag-b'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 1459, "Reservar", 0)&"</a>"
							end if
						else
							s_res = "<a class='link_serv' title='"&getTermo(global_idioma, 1407, "Efetue o login para reservar", 0)&"' style='cursor:pointer;' onclick=""abreLogin('reserva','"&getTermo(global_idioma, 348, "Reserva", 0)&"','&codigo_obra="&codigo_atual&"&tipo_obra="&Tipo&"&servidor="&iIndexSrv&"',false,true);""><span class='span_imagem div_imagem_right_3 icon_16 icon-small-flag-x'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 1459, "Reservar", 0)&"</a>"
						end if
						Response.Write "<tr><td class='esquerda td_ficha_serv' id='s_res"&codigo_atual&"' onMouseOver=servicosResultOver('s_res"&codigo_atual&"'); onMouseOut=servicosResultOut('s_res"&codigo_atual&"');>&nbsp;"&s_res&"</td></tr>"
					end if
					'*************************************
					'SUGESTÃO DE AQUISIÇÕES
					'*************************************
					if (serv_aquisicao = true) then
						if Session("Logado") = "sim" and Session("sgt_aquisicoes") = 1 and (Session("usuario_externo") = false) then
							if config_multi_servbib = 1 then
								if CStr(Session("Servidor_Logado")) = CStr(iIndexSrv) then
									s_aq = "<a class='link_serv' title='"&getTermo(global_idioma, 1408, "Sugerir aquisição", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/nova_sugestao.asp?codigo_obra="&codigo_atual&"&modo_busca=parent.hiddenFrame.modo_busca&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 1321, "Sugestões", 0)&"',710,425,false,false);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-lamp-b' data-icon='lamp-b'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 184, "Aquisição", 0)&"</a>"
									Response.Write "<tr><td class='esquerda td_ficha_serv' id='s_aq"&codigo_atual&"' onMouseOver=servicosResultOver('s_aq"&codigo_atual&"'); onMouseOut=servicosResultOut('s_aq"&codigo_atual&"');>&nbsp;"&s_aq&"</td></tr>"
								end if							
							else
								s_aq = "<a class='link_serv' title='"&getTermo(global_idioma, 1408, "Sugerir aquisição", 0)&"' style='cursor:pointer;' onclick=""abrePopup('asp/nova_sugestao.asp?codigo_obra="&codigo_atual&"&modo_busca=parent.hiddenFrame.modo_busca&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 1321, "Sugestões", 0)&"',710,425,false,false);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-lamp-b' data-icon='lamp-b'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 184, "Aquisição", 0)&"</a>"
								Response.Write "<tr><td class='esquerda td_ficha_serv' id='s_aq"&codigo_atual&"' onMouseOver=servicosResultOver('s_aq"&codigo_atual&"'); onMouseOut=servicosResultOut('s_aq"&codigo_atual&"');>&nbsp;"&s_aq&"</td></tr>"
							end if
						end if
					end if
					'*************************************
					'REFERÊNCIA
					'*************************************
					if (serv_referencia = true) then
						s_ref_info = "<a class='link_serv' title='"&serv_referencia_desc&"' style='cursor:pointer;' onclick=""abrePopup('asp/pop_referencia.asp?codigo="&codigo_atual&"&tipo="&Tipo&"&Servidor="&iIndexSrv&"&Idioma="&global_idioma&"','"&serv_referencia_desc&"',800,200,true,true);""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-book-b' data-icon='book-b'></span>&nbsp;&nbsp;"&serv_referencia_desc&"</a>"
						Response.Write "<tr><td class='esquerda td_ficha_serv' id='s_ref_info"&codigo_atual&"' onMouseOver=servicosResultOver('s_ref_info"&codigo_atual&"'); onMouseOut=servicosResultOut('s_ref_info"&codigo_atual&"');>&nbsp;"&s_ref_info&"</td>"
					end if
					'*************************************
					' RECIBO DO TITULO (ADEQUAÇÃO FAPCOM E BELAS ARTES)                 
					'*************************************
					if (global_numero_serie = 2160 or global_numero_serie = 617) then
						s_rec = "<a class='link_serv' title='Imprimir' style='cursor:pointer;' onclick=""LinkImpReciboTitulo("&iIndexSrv&","&codigo_atual&");""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-print' data-icon='print'></span>&nbsp;&nbsp;Imprimir</a>"
						Response.Write "<tr><td class='esquerda td_ficha_serv' id='s_rec"&codigo_atual&"' onMouseOver=servicosResultOver('s_rec"&codigo_atual&"'); onMouseOut=servicosResultOut('s_rec"&codigo_atual&"');>&nbsp;"&s_rec&"</td></tr>"						
					end if
					'*************************************
					'EXCLUIR REGISTRO DA MINHA SELEÇÂO
					'*************************************
					if (sOrigem = "MySel") then					
						s_einfo = "<a class='link_serv' title='"&getTermo(global_idioma, 2673, "Excluir de Minha seleção", 0)&"' style='cursor:pointer;' href='index.asp?modo_busca=parent.hiddenFrame.modo_busca&content=selecao&veio_de=menu&acao=excluir&codigo="&codigo_atual&"&tipo="&Tipo&"&Servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"'><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-delete-b' data-icon='delete-b'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 167, "Excluir", 0)&"</a>"
						Response.Write "<tr><td class='esquerda td_ficha_serv' id='s_einfo"&codigo_atual&"' onMouseOver=servicosResultOver('s_einfo"&codigo_atual&"'); onMouseOut=servicosResultOut('s_einfo"&codigo_atual&"');>&nbsp;"&s_einfo&"</td>"
					end if
					if (sOrigem = "favoritos") then					
						s_exclui = "<a class='link_serv' title='"&getTermo(global_idioma, 8348, "Excluir da lista de favoritos", 0)&"' style='cursor:pointer;' onclick=""excluir_titulo_favorito("&codigo_atual&", "&listaSelecionada&", "&iIndexSrv&");""><span class='transparent-icon span_imagem div_imagem_right_3 icon_16 icon-small-delete-b' data-icon='delete-b'></span>&nbsp;&nbsp;"&getTermo(global_idioma, 167, "Excluir", 0)&"</a>"
						Response.Write "<tr><td class='esquerda td_ficha_serv' id='s_rec"&codigo_atual&"' onMouseOver=servicosResultOver('s_rec"&codigo_atual&"'); onMouseOut=servicosResultOut('s_rec"&codigo_atual&"');>&nbsp;"&s_exclui&"</td></tr>"						
					end if
					
					Response.Write "</tr></table>"
					Response.Write "</td>"
				end if
			end if
			sequencial = sequencial + 1
		Next
	End if	
	Set xmlDoc = nothing
	Set xmlRoot = nothing		
End if

Response.Write "</table>"

%>
<!--#include file ="identificadoresDetalheImagem.asp"-->
<% 

'--------------------------------------------------------------------------------
' 				VERIFICA TWITTER HABILITADO
'--------------------------------------------------------------------------------
if global_twitter = 1 then %>
	<script>
        if(document.documentMode != 8) { 
            if (window.twttr == null) {      
                window.twttr=(function(d,s,id){var t,js,fjs=d.getElementsByTagName(s)[0];if(d.getElementById(id)){return}js=d.createElement(s);js.id=id;js.src="https://platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);return window.twttr||(t={_e:[],ready:function(f){t._e.push(f)}})}(document,"script","twitter-wjs"));
            } else {
                window.twttr.widgets.load();    
            }
        }
	</script>
<% end if %>
