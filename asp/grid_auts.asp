<%

if (left(sXMLFichas,5) = "<?xml") then
	'//--------------- Colunas do GRID ----------------------------------------------//
	Response.Write "<table class='max_width' style='border-spacing: 1px; padding: 0px'>"
	Response.Write "<tr>"
	Response.Write "<td class='td_tabelas_titulo centro' style='width: 35px; height: 20px;'>"
	if (busca_submetida <> "dsi") then
		Response.Write "&nbsp;#&nbsp;"
	else
		Response.Write "<input type='checkbox' id='cksel_todas' onclick='checarTodasAut();'>"
	end if
	Response.Write "</td>"
	Response.Write "<td class='td_tabelas_titulo centro' style='white-space: nowrap'>"&getTermo(global_idioma, 25, "Descrição", 0)&"</td>"
	Response.Write "<td class='td_tabelas_titulo centro' style='width: 90px'>&nbsp;"&getTermo(global_idioma, 83, "Tipo", 0)&"&nbsp;</td>"
	if (busca_submetida <> "dsi") then
		Response.Write "<td class='td_tabelas_titulo centro' style='width: 65px'>"&getTermo(global_idioma, 1197, "Pesquisar", 0)&"</td>"
	end if
	Response.Write "</tr>"

	sequencial = 1
	
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml sXMLFichas
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "Pagina" then
		For Each xmlFicha In xmlRoot.childNodes
			if xmlFicha.nodeName  = "Ficha" then
			
				if (sequencial mod 2) > 0 then '### IMPAR
					fontcolor = "black" 	
					td_class = "td_tabelas_valor2"
					link_class = "link_serv"
					css_search = "icon-small-search-w2"
                    css_pen = "icon-small-pen-w2"
				else '### PAR
					fontcolor= "#000000" 
					td_class = "td_tabelas_valor1"
					link_class = "link_serv"	
					css_search = "icon-small-search-w"
                    css_pen = "icon-small-pen-w"							
				end if				
			
				Registro = xmlFicha.attributes.getNamedItem("Registro").value
				
				Response.Write "<tr style='height: 20px'>"
				'//--------------------------------------------------------------------//
				'//------------------- NUMERO DE ORDEM NA PROCURA ---------------------//
				'//--------------------------------------------------------------------//			
				codigo_atual = 0
				For Each xmlCampos In xmlFicha.childNodes
					'*************************************
					'CODIGO
					'*************************************
					if xmlCampos.nodeName = "CODIGO" then
						codigo_atual = xmlCampos.attributes.getNamedItem("Valor").value
						
						'------------- Indicador de posição na busca ------
						Response.write "<td class='centro "&td_class&"'>"
						if (busca_submetida <> "dsi") then				
							Response.write "<span style='color: black'><b>&nbsp;"&Registro&"&nbsp;</b></span>"
						else
							Response.write "<input type='checkbox' id='cksel_aut"&codigo_atual&"' value='"&codigo_atual&"'>"
						end if
						Response.write "</td>"
					end if
					'*************************************
					'TITULO DA AUTORIDADE
					'*************************************
					if xmlCampos.nodeName = "TITULO" then
						descricao_aut = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("Valor").value)
						desc_princ_aut = xmlCampos.attributes.getNamedItem("Desc_Princ").value						
						tipo_aut = xmlCampos.attributes.getNamedItem("Tipo").value
						autor_formatado = replace(replace(replace(replace(desc_princ_aut," ","_"),"<",""),">",""),"'","\'")
						'*************************************
						'DESCRICAO
						'*************************************
						descricao = "<td class='esquerda "&td_class&"'>"
						inicio_dec = Registro - 1
	
						if ((global_marc = 1) or (global_recursos_avancados = 1)) and (busca_submetida <> "dsi") then
							descricao = descricao & "<a class='" & link_class & "' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalheAutoridade(parent.hiddenFrame.modo_busca,"&GetSessao(global_cookie,"nrows_auts"&iIndexSrv)&","&inicio_dec&","&codigo_atual&","&pagina&",'resultado'," &tipo_aut & ");"">&nbsp;" & descricao_aut & "</a>"
						else
							descricao = descricao &  "&nbsp;" & descricao_aut 
						end if
						Response.write descricao&"</td>"
						'*************************************
						'TIPO
						'*************************************
						Select case tipo_aut
							case "100"
								if (busca_submetida <> "dsi") then
									imagem = getTermo(global_idioma, 732, "Pessoa", 0)
								else
									imagem = getTermo(global_idioma, 61, "Autor", 0)
								end if
							case "110"
								imagem = getTermo(global_idioma, 62, "Instituição", 0)
							case "111"
								imagem = getTermo(global_idioma, 63, "Evento", 0)
							case "130"
								imagem = getTermo(global_idioma, 70, "Título uniforme", 0)
							case "148"
								imagem = getTermo(global_idioma, 71, "Termo cronológico", 0)
							case "150"
								if (busca_submetida <> "dsi") then
									imagem = getTermo(global_idioma, 742, "Termo tópico", 0)
								else
									imagem = getTermo(global_idioma, 72, "Assunto", 0)
								end if
							case "151"
								imagem = getTermo(global_idioma, 1148, "Local geográfico", 0)
							case "155"
								imagem = getTermo(global_idioma, 3707, "Termo de gênero e forma", 0)
							case "180"
								imagem = getTermo(global_idioma, 75, "Subdivisão geral", 0)
							case "181"
								imagem = getTermo(global_idioma, 76, "Subdivisão geográfica", 0)
							case "182"
								imagem = getTermo(global_idioma, 77, "Subdivisão cronológica", 0)
							case "185"
								imagem = getTermo(global_idioma, 3390, "Subdivisão de forma", 0)
							case else
								imagem = "else"&tipo_aut
						End Select
						Response.write "<td class='"&td_class&" centro'>"&Imagem&"</td>"						
						'*************************************
						'LINK
						'*************************************
						if (busca_submetida <> "dsi") then
							Select case tipo_aut
								case "100"
	                                imagem = "<a class= centro title='"&getTermo(global_idioma, 2147, "Pesquisar todos os registros relacionados a este autor", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"								
   								case "110"
	                                imagem = "<a class= centro title='"&getTermo(global_idioma, 2656, "Pesquisar todos os registros relacionados a esta instituição", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"																
   								case "111"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 2657, "Pesquisar todos os registros relacionados a este evento", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
   								case "130"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 2658, "Pesquisar todos os registros relacionados a este título uniforme", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
    							case "148"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 7484, "Pesquisar todos os registros relacionados a este termo cronológico", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
    							case "150"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 2659, "Pesquisar todos os registros relacionados a este termo tópico", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
    							case "151"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 7485, "Pesquisar todos os registros relacionados a este local geográfico", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
    							case "155"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 7486, "Pesquisar todos os registros relacionados a este termo de gênero e forma", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
    							case "180"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 7487, "Pesquisar todos os registros relacionados a esta subdivisão geral", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
    							case "181"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 7488, "Pesquisar todos os registros relacionados a esta subdivisão geográfica", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
    							case "182"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 7489, "Pesquisar todos os registros relacionados a esta subdivisão cronológica", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
    							case "185"
									imagem = "<a class= centro title='"&getTermo(global_idioma, 7490, "Pesquisar todos os registros relacionados a esta subdivisão de forma", 0)&"...' style='cursor:pointer;' onClick=LinkBuscaAutoridade('rapida',"&codigo_atual&",'"&autor_formatado&"',"&iIndexSrv&") ><span class='transparent-icon span_imagem icon_16 "&css_search&"' data-icon='search'></span></a>"
    							case else
	                                imagem = "<span class='transparent-icon span_imagem icon_16 "&css_pen&"' data-icon='pen'></span>"
							End Select
							Response.write "<td class='centro "&td_class&"'>"&Imagem&"</td>"
						end if
					end if				
				Next
				
				Response.write "</tr>"  '-> Fecha a linha
			end if
			sequencial = sequencial + 1
		Next
	end if
	Set xmlDoc = nothing
	Set xmlRoot = nothing
	
	Response.Write "<tr>"
	Response.Write "<td colspan='4'>&nbsp;</td>"
	Response.Write "</tr>"
	
	Response.write "</table>"
end if
	
%>