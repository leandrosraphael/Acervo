<%
function ExibirTagSubCampos(XML)
	SubCampo = XML.attributes.getNamedItem("SubCmp").value
	if (XML.nodeName  = "Reg") and _
		(	(SubCampo <> "a") and (SubCampo <> "b") and (SubCampo <> "c") and _
			(SubCampo <> "d") and (SubCampo <> "t") and (SubCampo <> "x") and _
			(SubCampo <> "v") and (SubCampo <> "y") and (SubCampo <> "z")) then
		ExibirTagSubCampos = true        
	else
		ExibirTagSubCampos = false
	end if
end function

Function AdicionaEspacoInicio(valor)
	AdicionaEspacoInicio = ""
	for i = 1 to len(valor)
		if (mid(valor, i, 1) = " ") then
			AdicionaEspacoInicio = AdicionaEspacoInicio & "&nbsp;"
		else
			AdicionaEspacoInicio = AdicionaEspacoInicio & right(valor, len(valor) - i + 1)
			exit for
		end if
	next
End Function

Function AdicionaEspaco(valor)
	AdicionaEspaco = replace(valor," ","&nbsp;")
End Function

Function ObterDescricaoCampo(campo)
    descricao = ""

    if (campo = "LC") then
        descricao = "Library of Congress"
    elseif (campo = "MESH") then
        descricao = "MeSH - Medical Subject Headings"
    elseif (campo = "NAL") then
        descricao = "NAL - National Agricultural Library"
    elseif (campo = "NE") then        
        descricao = getTermo(global_idioma, 1967, "Não especificada", 0)
	elseif (campo = "CSH") then
        descricao = "CSH - Canadian Subject Headings"
	elseif (campo = "RVM") then
        descricao = "RVM - Répertoire de vedettes-matière"
	elseif (campo = "OF") then
        descricao = getTermo(global_idioma, 2653, "Outras fontes", 0)
    else
        descricao = campo
    end if

    ObterDescricaoCampo = descricao
End Function 

Function IsCampoOutroVocabulario(campo)
    IsCampoOutroVocabulario = (campo = "LC") OR (campo = "MESH") OR (campo = "NAL") OR (campo = "NE") OR (campo = "CSH") OR (campo = "RVM") OR (campo = "OF")
End Function

Function Formata_Ficha(stringXML,modo,idioma, servidor)
    ValorTags110 = ""
    ValorTagsRepet = ""
    identificadoresInfo = ""
   
	if modo = "obra" then
		css_desc = "td_marc_descricao"
		css_desc2 = "td_marc_descricao_recuo"
	else
		css_desc = "td_detalhe_descricao"
		css_desc2 = "td_detalhe_descricao_recuo"
	end if
	conta_campos = 0
	if left(stringXML,5) <> "<?xml" then
		Formata_Ficha = ""
	else
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml stringXML
		Set xmlRoot = xmlDoc.documentElement

		strSubGeo = ""

		if (Not xmlRoot.attributes.getNamedItem("Tipo") Is Nothing) then
			strDoc = strDoc & "<!--tipo:"&xmlRoot.attributes.getNamedItem("Tipo").value&"/tipo-->"
		end if

		if (Not xmlRoot.attributes.getNamedItem("NaoRevisado") Is Nothing) then
			strDoc = strDoc & "<!--nao_revisado-->"
		end if

		if (Not xmlRoot.attributes.getNamedItem("SubGeo") Is Nothing) then
			strSubGeo = " <label class='label-subgeo'>(subdividido geograficamente)</label>"
		end if

		strDoc = strDoc & "<table class='max_width table-ficha-detalhes' style='border-spacing: 1px; padding: 0'>"
    
		if modo = "auts" then
            possuiCampoOutroVocabulario = false

			Dim campos
			Set campos = CreateObject("Scripting.Dictionary")

			contLink = 1
			For Each xmlPNode In xmlRoot.childNodes
				If xmlRoot.childNodes.length = 0 Then
					strDoc = ""
				Else
                      
                    if xmlPNode.nodeName  = "IDENTIFICADORES" then  
                        For Each xmlIdentificadores in xmlPNode.childNodes
                            tipoIdentificador = xmlIdentificadores.attributes.getNamedItem("Tipo").value
                            imagemIdentificador = "data:image;base64," & xmlIdentificadores.attributes.getNamedItem("Base64").value
                            strdoc = strdoc + "<script>sessionStorage.setItem('"&tipoIdentificador&"', '"&imagemIdentificador&"');</script>"                    
                        next
                    end if
                     
					if xmlPNode.nodeName  = "Reg" then                        
						conta_campos = conta_campos + 1
						Tag = xmlPNode.attributes.getNamedItem("Tag").value
						Desc = trim(xmlPNode.attributes.getNamedItem("Desc").value)
						Valor = xmlPNode.attributes.getNamedItem("Valor").value
						Nivel = trim(xmlPNode.attributes.getNamedItem("Nivel").value)

						Codigo = 0
						Busca = ""
						CodigoBusca = 0
						Tipo = 0
						Sigla = ""
						Titulo = False
						if (Not xmlPNode.attributes.getNamedItem("Codigo") Is Nothing) then
							Codigo = xmlPNode.attributes.getNamedItem("Codigo").value
						end if
						if (Not xmlPNode.attributes.getNamedItem("Busca") Is Nothing) then
							Busca = xmlPNode.attributes.getNamedItem("Busca").value
						end if
						if (Not xmlPNode.attributes.getNamedItem("CodigoBusca") Is Nothing) then
							CodigoBusca = xmlPNode.attributes.getNamedItem("CodigoBusca").value
						end if
						if (Not xmlPNode.attributes.getNamedItem("Tipo") Is Nothing) then
							Tipo = xmlPNode.attributes.getNamedItem("Tipo").value
						end if
						if (Not xmlPNode.attributes.getNamedItem("Sigla") Is Nothing) then
							Sigla = "<b>" & xmlPNode.attributes.getNamedItem("Sigla").value & "</b>&nbsp;"
						end if
						if (Not xmlPNode.attributes.getNamedItem("Titulo") Is Nothing) then
							Titulo = True
						end if

						if (Tag = "LATTES") then
							Valor = "<span class='span_imagem icon_16 icon-lattes' style='vertical-align: -4px;'></span>&nbsp;<a href='"&Valor&"' target='_blank'>" & Valor & "</a>"
						elseif (Tag = "7xx.01") then
							Valor = "<span class='span_imagem icon_16 icon-loc' style='vertical-align: -4px;'></span>&nbsp;" & Valor
                            possuiCampoOutroVocabulario = true
						elseif (Tag = "7xx.2") then
							Valor = "<span class='span_imagem icon_16 icon-mesh' style='vertical-align: -4px;'></span>&nbsp;" & Valor
                            possuiCampoOutroVocabulario = true
						elseif (Tag = "7xx.3") then
							Valor = "<span class='span_imagem icon_16 icon-nal' style='vertical-align: -4px;'></span>&nbsp;" & Valor
                            possuiCampoOutroVocabulario = true
						elseif (Tag = "7xx.5") then
							Valor = "<span class='span_imagem icon_16 icon-csh' style='vertical-align: -4px;'></span>&nbsp;" & Valor
                            possuiCampoOutroVocabulario = true
						elseif (Tag = "7xx.6") then
							Valor = "<span class='span_imagem icon_16 icon-rvm' style='vertical-align: -4px; width: 40px;'></span>&nbsp;" & Valor
                            possuiCampoOutroVocabulario = true
                        elseif (Tag = "7xx.4") OR (Tag = "7xx.7") then
                            possuiCampoOutroVocabulario = true
                        elseif (Tag = "024") then
 
		                        Set xmlIdentificadores = CreateObject("Microsoft.xmldom")
		                        xmlIdentificadores.async = False
                                set xmlIdentificadores = xmlPNode

                                %><!--#include file ="../asp/identificadoresDetalhe.asp"--><%

                                Desc = ""
                                Valor = ""
                                Nivel = ""

                        elseif (Tag = "100") then
                            Valor = TrocaTagMarcador(Valor)

                            if identificadoresInfo <> "" then
                                Valor = Valor + "&nbsp;" + identificadoresInfo   
                            end if
						else
                            Valor = TrocaTagMarcador(Valor)
                            Valor = replace(replace(replace(replace(Valor,"#URL#","<a target='_blank' "), "#HREF#", "href='"), "#/HREF#", "'>"), "#/URL#", "</a>")
						end if    
                            
						if (Tag = "675") or (Tag = "682") then                             
							For Each xmlSubCampos In xmlPNode.childNodes
                                if (TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value) <> xmlPNode.attributes.getNamedItem("Valor").value) then
                                    ValorTagsRepet = ValorTagsRepet & "<tr>"                                    
                                    ValorTagsRepet = ValorTagsRepet & "<td class='"&css_desc2&" esquerda'>"&Desc&"&nbsp;</td>"
                                    ValorTagsRepet = ValorTagsRepet & "<td class='td_detalhe_valor'><div class='justificado'>" & TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value)  & "&nbsp;</div></td>"
                                    ValorTagsRepet = ValorTagsRepet & "</tr>"                 							
                                end if 
							next
                            if (ValorTagsRepet <> "") then
                                Valor = Valor & ValorTagsRepet
                                ValorTagsRepet = ""
                            end if
						end if
    
						if (Busca <> "") and (Tag <> "024") then
							ValorFormatado = replace(replace(replace(replace(Busca," ","_"),"<",""),">",""),"'","\'")

							ValorHtml = "<table class='autLink' style='border-spacing: 1px; padding: 0;'>"

							ValorHtml = ValorHtml & "<tbody>"
							ValorHtml = ValorHtml & "<tr>"
							ValorHtml = ValorHtml & "<td id='autLinkTab1' class='td_left_middle'>"
							ValorHtml = ValorHtml & Sigla

							select case Tipo
								case "100"
									tipo_img = "autImg"
									hint_busca = getTermo(global_idioma, 1558, "Buscar todos os registros deste autor", 0)&"..."
									hint_link = getTermo(global_idioma, 1571, "Mostrar informações sobre este autor", 0)
								case "110"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 1559, "Buscar todos os registros desta instituição", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "111"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 1560, "Buscar todos os registros deste evento", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "130"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 1561, "Buscar todos os registros deste título unificado", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "148"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 7491, "Buscar todos os registros deste termo cronológico", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "150"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 1562, "Buscar todos os registros deste termo tópico", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "151"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 7492, "Buscar todos os registros deste local geográfico", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "155"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 7493, "Buscar todos os registros deste termo de gênero e forma", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "180"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 7494, "Buscar todos os registros desta subdivisão geral", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "181"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 7495, "Buscar todos os registros desta subdivisão geográfica", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "182"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 7496, "Buscar todos os registros desta subdivisão cronológica", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case "185"
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 7497, "Buscar todos os registros desta subdivisão de forma", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)
								case else
									tipo_img = "assuntoImg"
									hint_busca = getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"..."
									hint_link = getTermo(global_idioma, 1572, "Mostrar informações sobre este assunto", 0)&"..."
							End Select

							If Titulo then
								ValorHtml = ValorHtml & "<b>"
							end if                           
                            
							ValorHtml = ValorHtml & "<a class='link_classic2' title='" & hint_busca & "' href='javascript:LinkBuscaAutoridade(parent.hiddenFrame.modo_busca," & CodigoBusca & ",""" & ValorFormatado & """," & servidor & ");'>" & Valor & "</a>&nbsp"
							
							If Titulo then
								ValorHtml = ValorHtml & "</b>" & strSubGeo
							end if

							if Codigo > 0 then
								ValorHtml = ValorHtml & "</td>"
								ValorHtml = ValorHtml & "<td class='td_left_middle "&tipo_img&"' id='"&tipo_img & contLink & "' onclick='LinkAutInfo(""detalhes""," & Codigo & ",""" & ValorFormatado & ""","&tipo_assunto&")' style='cursor:pointer;' title='"&hint_link&"'>"
							end if

							ValorHtml = ValorHtml & "</td>"
							ValorHtml = ValorHtml & "</tr>"
        						
							if (Tag = "110") and (global_numero_serie = 5592) then                       
								For Each xmlSubCampos In xmlPNode.childNodes
									if (ExibirTagSubcampos(xmlSubCampos)) then                        
										ValorTags110 = ValorTags110 & "<tr>"
										ValorTags110 = ValorTags110 &		"<td class='td_detalhe_descricao_imagem'>"
										ValorTags110 = ValorTags110 &			"<div class='joinbottom icon_20'></div>"
										ValorTags110 = ValorTags110 &		"</td>"
										ValorTags110 = ValorTags110 &		"<td class='td_detalhe_descricao_recuo'>"& xmlSubCampos.attributes.getNamedItem("Desc").value & "</td>"
										ValorTags110 = ValorTags110 &		"<td class='td_detalhe_valor'>"
										ValorTags110 = ValorTags110 &			"<div class='justificado'>"& TrocaTagMarcador(xmlSubCampos.attributes.getNamedItem("Valor").value) & "</div>"
										ValorTags110 = ValorTags110 &		"</td>"
										ValorTags110 = ValorTags110 & "</tr>"                        
									end if
								next
							end if
							ValorHtml = ValorHtml & "</tbody>"
							ValorHtml = ValorHtml & "</table>"
						
							contLink = contLink + 1
						else
							ValorHtml = Valor
						end if
						
                        if ValorHtml <> "" then
						    if (not campos.Exists(Desc)) then
							    campos.Add Desc, Array()
						    end if

						    dim aTemp 
						    aTemp = campos(Desc)
						    novo = ubound(aTemp) + 1
						    ReDim Preserve aTemp(novo)
						    aTemp(novo) = ValorHtml

						    campos(Desc) = aTemp
                        end if 

					end if	                    
				end if
			next	

			strLinhas = ""
			strLC = ""
            strMeSH = ""
            strNAL = ""
            strCSH = ""
            strNaoEspecificada = ""
            strOutrasFontes = ""
            strRVM = ""
			strDesc = ""
			AdicionarInformacoesInstituicao = (ValorTags110 <> "")
			For Each campo In campos.Keys
	
                if (IsCampoOutroVocabulario(campo)) then
                    colunaRecuo = "<td class='td_detalhe_descricao_imagem'><div class='joinbottom icon_20'></div></td>"
                    colSpan = ""
                    classCss = "class='direita td_detalhe_descricao_recuo'"
                else
                    colunaRecuo = ""
                    
                    if (possuiCampoOutroVocabulario) then
                        colSpan = "colspan='2'"
                    end if

                    classCss = "class='direita td_detalhe_descricao'"
                end if            

				strLinha = "<tr>"
                strLinha = strLinha & colunaRecuo
				if (AdicionarInformacoesInstituicao)  then
					strLinha = strLinha & "<td colspan='2' " & classCss & ">"
				else
					strLinha = strLinha & "<td " & colSpan & " " & classCss & ">"
				end if
				strLinha = strLinha & ObterDescricaoCampo(campo)
				strLinha = strLinha & "</td>"

				if (AdicionarInformacoesInstituicao) then
				    if (identificadoresInfo <> "") then
					    strLinha = strLinha & "<td colspan='2' class='td_detalhe_valor line-height-0'>"
                    else
					    strLinha = strLinha & "<td colspan='2' class='td_detalhe_valor'>"
                    end if
				else
				    if (identificadoresInfo <> "") then
					    strLinha = strLinha & "<td class='td_detalhe_valor line-height-0'>"
				    else 
					    strLinha = strLinha & "<td class='td_detalhe_valor'>"
				    end if
				end if

				identificadoresInfo = ""
				
				lista = ((ubound(campos(campo)) > 0) and (InStr(campos(campo)(0), "<table") = 0))

				if lista then
					strLinha = strLinha & "<ul>"
				end if
	
				for i = 0 to ubound(campos(campo))
					if lista then
						strLinha = strLinha & "<li style='list-style: none;'>"
					end if

					strLinha = strLinha & campos(campo)(i)
	
					if lista then
						strLinha = strLinha & "</li>"
					end if
				next 
	
				if lista then
					strLinha = strLinha & "</ul>"
				end if 	

				strLinha = strLinha & "</td>"
				strLinha = strLinha & "</tr>"

				if (ValorTags110 <> "") then
					strLinha = strLinha & ValorTags110
					ValorTags110 = ""
				end if

				if (campo = "LC") then
					strLC = strLinha
				elseif (campo = "MESH") then
                    strMeSH = strLinha                                        
				elseif (campo = "NAL") then
                    strNAL = strLinha
				elseif (campo = "NE") then
                    strNaoEspecificada = strLinha
				elseif (campo = "CSH") then
                    strCSH = strLinha
				elseif (campo = "RVM") then
                    strRVM = strLinha
				elseif (campo = "OF") then
                    strOutrasFontes = strLinha
                else
					if strDesc = "" then
						strDesc = strLinha
					else
						strLinhas = strLinhas & strLinha
					end if
				end if
			Next	

            if (possuiCampoOutroVocabulario) then
                strCabecalhoOutroVocabulario = "<tr><td colspan='2' class='direita td_detalhe_descricao'>" & getTermo(global_idioma, 8216, "Outros vocabulários", 0) & "</td><td class='td_detalhe_valor'></td></tr>"
            else
                strCabecalhoOutroVocabulario = ""
            end if

			strDoc = strDoc & strDesc & strLinhas & strCabecalhoOutroVocabulario & strLC & strMeSH & strNAL & strNaoEspecificada & strCSH & strRVM & strOutrasFontes 
		else

			For Each xmlPNode In xmlRoot.childNodes
				If xmlRoot.childNodes.length = 0 Then
					strDoc = ""
				Else

					if xmlPNode.nodeName  = "Reg" then
						conta_campos = conta_campos + 1
						Desc = trim(xmlPNode.attributes.getNamedItem("Desc").value)
						Valor = xmlPNode.attributes.getNamedItem("Valor").value
						strDoc = strDoc & "<tr>"
						strDoc = strDoc & "<td colspan='2' class='"&css_desc&" esquerda'>"&replace(replace(replace(Desc,chr(10),"<br />"),"\","\\"),"""","\""")&"&nbsp;</td>"
						strDoc = strDoc & "<td class='td_detalhe_valor'>" & TrocaTagMarcador(replace(replace(replace(Valor,chr(10),"<br />"),"\","\\"),"""","\""")) & "&nbsp;</td>"
						strDoc = strDoc & "</tr>"
					end if			
					For Each xmlNode In xmlPNode.childNodes
						if xmlNode.nodeName  = "Reg" then
 							conta_campos = conta_campos + 1
							SubCmp = xmlNode.attributes.getNamedItem("SubCmp").value
							Desc = trim(xmlNode.attributes.getNamedItem("Desc").value)
							Valor = xmlNode.attributes.getNamedItem("Valor").value
							Valor = TrocaTagMarcador(replace(replace(replace(Valor,chr(10),"<br />"),"\","\\"),"""","\"""))
							'Se for tag 856\u adiciona um link
							if ((Tag = "856") AND (lCase(SubCmp) = "u")) then
								Valor = "<a class='link_classic2' href='" & Valor & "' target='_blank'>" & Valor & "</a>"
							end if

							strDoc = strDoc & "<tr>"
							strDoc = strDoc & "<td class='centro' style='height: 16px; width: 19px'><div class='joinbottom icon_20'></div></td>"
							strDoc = strDoc & "<td class='"&css_desc2&" esquerda'>"&Desc&"&nbsp;</td>"
							strDoc = strDoc & "<td class='td_detalhe_valor'><div class='justificado'>" & Valor & "&nbsp;</div></td>"
							strDoc = strDoc & "</tr>"
						end if
					next	
	
				End If
			Next
		end if

        %><!--#include file ="../asp/identificadoresDetalheImagem.asp"--><%
		strDoc = strDoc & "</table>"

		if conta_campos > 0 then
			Formata_Ficha = strDoc
		else
			Formata_Ficha = getTermo(idioma, 1273, "Nenhuma informação encontrada", 0)
		end if
		Set xmlDoc = nothing
		Set xmlRoot = nothing
	End If
End Function

Function Formata_Tags(stringXML,modo,idioma,codigo,tipo,servidor)
	if modo = "obra" then
		css_desc = "#E8F2FD"
	else
		css_desc = "#E8F2FD"
	end if
	conta_campos = 0

    strDoc = ""
	
    if left(stringXML,5) <> "<?xml" then
		Formata_Tags = ""
	else
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml stringXML
		Set xmlRoot = xmlDoc.documentElement
		
		For Each xmlPNode In xmlRoot.childNodes
			If xmlRoot.childNodes.length = 0 Then
				strDoc = ""
			Else
				if xmlPNode.nodeName  = "Reg" then
					conta_campos = conta_campos + 1					
				
					Tag = trim(xmlPNode.attributes.getNamedItem("Tag").value)
					Ind = xmlPNode.attributes.getNamedItem("Ind").value
					Desc = trim(xmlPNode.attributes.getNamedItem("Desc").value)
					Valor = xmlPNode.attributes.getNamedItem("Valor").value
					Valor = AdicionaEspaco(replace(replace(replace(Valor,chr(10),"<br />"),"\","\\"),"""","\"""))
											
                    strDoc = strDoc & "<b>" & Tag & "</b>"
                    if (CInt(Trim(Tag)) > 8) then
                        'Formatando os indicadores
                        strDoc = strDoc & " "
                        if (Trim(Ind) = "") then
                            strDoc = strDoc & "__"
                        else
                            if (Len(CStr(Ind)) > 1) then
                                Ind1 = Trim(Mid(Ind, 1, 1))
                                Ind2 = Trim(Mid(Ind, 2, 1))
                            else
                                Ind1 = Trim(Ind)
                                Ind2 = ""
                            end if

                            Indicador = ""
                            if (Ind1 <> "") then
                                if (Ind2 <> "") then
                                    Indicador = Ind1 & Ind2
                                else
                                    Indicador = Ind1 & "_"
                                end if
                            elseif (Ind2 <> "") then
                                Indicador = "_" & Ind2
                            end if

                            strDoc = strDoc & CStr(Indicador)
                        end if
                    end if
				end if			
				For Each xmlNode In xmlPNode.childNodes
					if xmlNode.nodeName  = "Reg" then
 						conta_campos = conta_campos + 1					
								
						SubCmp = trim(xmlNode.attributes.getNamedItem("Tag").value)
						Desc = trim(xmlNode.attributes.getNamedItem("Desc").value)
						Valor = TrocaTagMarcador(xmlNode.attributes.getNamedItem("Valor").value)
						
						Valor = replace(replace(replace(Valor,chr(10),"<br />"),"\","\\"),"""","\""")

						if (Tag <= "008") then
							Valor = AdicionaEspaco(Valor)
						else
							Valor = AdicionaEspacoInicio(Valor)
						end if

						'Se for tag 856\u adiciona um link
						if ((modo = "obra") AND (Tag = "856") AND (lCase(SubCmp) = "u")) then
							Plataforma = xmlNode.attributes.getNamedItem("Plataforma").value
							CodMidia = xmlNode.attributes.getNamedItem("Codigo").value
							if (Plataforma = "") then
								Valor = "<a class='link_classic2' href='" & Valor & "' target='_blank'>" & Valor & "</a>"
							else
								Valor = "<a class='link_classic2' href='javascript:abreMidiaEspecifica(" & tipo & "," & codigo & "," & CodMidia & "," & servidor & ")'>" & Valor & "</a>"
							end if
						end if

                        if (Trim(SubCmp) = "") then
                            strDoc = strDoc & " " & Valor
                        else
                            'Formatando os subcampos
                            strDoc = strDoc & " <b>|" & SubCmp & "</b>"

                            'Formatando os valores
                            strDoc = strDoc & " " & Valor
                        end if
					end if
				next				
                strDoc = strDoc & "<br />"
			End If			
		Next	
		if conta_campos > 0 then
			Formata_Tags = strDoc
		else
			Formata_Tags = getTermo(idioma, 1273, "Nenhuma informação encontrada", 0)
		end if
		Set xmlDoc = nothing
		Set xmlRoot = nothing
	End If
End Function
%>
