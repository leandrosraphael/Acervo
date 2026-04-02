<%
'--------------------------------------------------------------------------------
' 				MONTA RESULTADO EDS EM FORMATO DE FICHA
'--------------------------------------------------------------------------------
Response.Write "<table class='max_width' style='border-spacing: 1px; padding: 0'><tr>"
pagAtualEds = 1
if Request.QueryString("pagina_eds") <> "" then
	pagAtualEds = CInt(Request.QueryString("pagina_eds"))
end if
sLinhas = GetSessao(global_cookie,"linhas_eds")
numLinhasEds = 10
if sLinhas <> null and sLinhas <> "" then
	numLinhasEds = CInt(sLinhas)
end if

if (left(sXMLFichasEds,5) = "<?xml") then
	contRegistro = 1 + ((pagAtualEds - 1) * numLinhasEds)
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml sXMLFichasEds
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "SearchResult" then
		For Each xmlSearchResult In xmlRoot.childNodes
			if xmlSearchResult.nodeName = "Records" then
				For Each xmlFichaRecords In xmlSearchResult.childNodes
					
					if xmlFichaRecords.nodeName = "Record" then
						ficha = ""
						links = ""
						sDbId = ""
						sAn = ""
						For Each xmlFichaRecord In xmlFichaRecords.childNodes
							if xmlFichaRecord.nodeName = "Header" then
								sPubTypeId = ""
								sPubType = ""

								For Each xmlFichaHeader In xmlFichaRecord.childNodes
									if xmlFichaHeader.nodeName = "PubType" then
										sPubType = xmlFichaHeader.text
									elseif xmlFichaHeader.nodeName = "PubTypeId" then
										sPubTypeId = xmlFichaHeader.text
									elseif xmlFichaHeader.nodeName = "DbId" then
										sDbId = xmlFichaHeader.text
									elseif xmlFichaHeader.nodeName = "An" then
										sAn = xmlFichaHeader.text
									end if
								Next

								Response.Write "<tr>"
								'//-------------------------------------------------------------//
								'---------------------CELULA COM NUMERO ---------------//
								'//-------------------------------------------------------------//
						
								Response.Write "<td class='td_center_top td_grid_ficha_background' style='width:40px; white-space: nowrap'>"
						
								Response.Write "<br /><table style='display: inline-block'><tr><td class='centro' style='font-size: 13;'>"
								Response.Write contRegistro&"<br />"
								Response.Write "</td></tr></table>"
						
								Response.Write "</td>"
						
								'//-------------------------------------------------------------//
								'---------------------CELULA COM A FICHA------------------------//
								'//-------------------------------------------------------------//
						
								
								Response.Write "<td class='td_center_top td_grid_ficha_background' style='width: 120px; white-space: nowrap'>"
								if sPubTypeId <> "" then
									Response.Write "<table style='display: inline-table'><tr><td>"
									Response.Write "<br/>"
									Response.Write "<span class='pt-icon pt-"&sPubTypeId&"'></span>"
									Response.Write "<br/>"
									Response.Write sPubType
									Response.Write "</td></tr></table>"
								end if
								Response.Write "</td>"
						
								Response.Write "<td class='td_center_top td_grid_ficha_background'>"
						
						
							end if

							if xmlFichaRecord.nodeName = "Html" then
								if xmlFichaRecord.text = "1" then
									if(Session("Logado")= "sim") then
										links = links & "<tr><td class='esquerda td_ficha_serv' id='eds_html"&contRegistro&"' onMouseOver=servicosResultOver('eds_html"&contRegistro&"'); onMouseOut=servicosResultOut('eds_html"&contRegistro&"');>&nbsp;<a class='link_serv' title='Full text (HTML)' style='cursor:pointer;' onclick=""abrePopup('asp/eds_full_text.asp?type=0&dbid="&sDbId&"&an="&sAn&"','Full Text (HTML)',800,600,true,true,null,true);""><span class='span_imagem icon_11 icon-txt'></span>&nbsp;&nbsp;HTML</a></td></tr>"
									else
										links = links & "<tr><td class='esquerda td_ficha_serv' id='eds_html"&contRegistro&"' onMouseOver=servicosResultOver('eds_html"&contRegistro&"'); onMouseOut=servicosResultOut('eds_html"&contRegistro&"');>&nbsp;<a class='link_serv' title='Full text (HTML)' style='cursor:pointer;' onclick=""abreLogin('eds_full_text', 'Full Text (HTML)', '&pagina="&pagAtualEds&"', false, true);""><span class='span_imagem icon_11 icon-txt'></span>&nbsp;&nbsp;HTML</a></td></tr>"
									end if
								end if
							end if

							if xmlFichaRecord.nodeName = "Pdf" then
								if xmlFichaRecord.text = "1" then
									if(Session("Logado")= "sim") then
										links = links & "<tr><td class='esquerda td_ficha_serv' id='eds_pdf"&contRegistro&"' onMouseOver=servicosResultOver('eds_pdf"&contRegistro&"'); onMouseOut=servicosResultOut('eds_pdf"&contRegistro&"');>&nbsp;<a class='link_serv' title='Full text (PDF)' style='cursor:pointer;' onclick=""abrePopup('asp/eds_full_text.asp?type=1&dbid="&sDbId&"&an="&sAn&"','Full Text (PDF)',800,600,true,true,null,true);""><span class='span_imagem icon_11 icon-pdf'></span>&nbsp;&nbsp;PDF</a></td></tr>"
									else
										links = links & "<tr><td class='esquerda td_ficha_serv' id='eds_pdf"&contRegistro&"' onMouseOver=servicosResultOver('eds_pdf"&contRegistro&"'); onMouseOut=servicosResultOut('eds_pdf"&contRegistro&"');>&nbsp;<a class='link_serv' title='Full text (PDF)' style='cursor:pointer;' onclick=""abreLogin('eds_full_text', 'Full Text (HTML)', '&pagina="&pagAtualEds&"', false, true);""><span class='span_imagem icon_11 icon-pdf'></span>&nbsp;&nbsp;PDF</a></td></tr>"
									end if
								end if
							end if
						
							if xmlFichaRecord.nodeName = "Items" then
								For Each xmlFichaItems In xmlFichaRecord.childNodes
									if xmlFichaItems.nodeName = "Item" then
										sLabel = ""
										sData = ""
										For Each xmlFichaItem In xmlFichaItems.childNodes
											if xmlFichaItem.nodeName = "Label" then
												sLabel = xmlFichaItem.text
											elseif xmlFichaItem.nodeName = "Data" then
												sData = xmlFichaItem.text
												sData = Replace(sData, "/highlight", "/span")
												sData = Replace(sData, "highlight", "span class='destaca_palavras'")
                                                if (InStr(sData, "<link") > 0) then
                                                    sData = Replace(sData, "linkWindow=", "target=")
                                                    sData = Replace(sData, "linkTerm=", "href=")
                                                    sData = Replace(sData, "</link", "</a")
												    sData = Replace(sData, "<link", "<a")
                                                end if
                                                if (InStr(sData, "<externalLink") > 0) then
                                                    sData = Replace(sData, "</externalLink", "</a")
												    sData = Replace(sData, "<externalLink", "<a target='_blank'")
                                                    sData = Replace(sData, "term=", "href=")
                                                end if
											end if
										Next
							
										ficha = ficha & "<tr>"
										ficha = ficha & "<td class='td_ficha_esq esquerda'>"
										ficha = ficha & sLabel
										ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&sData&"</td></tr>"
							
									end if
								Next
							end if

							if xmlFichaRecord.nodeName = "Links" then
								For Each xmlFichaLinks In xmlFichaRecord.childNodes
									if xmlFichaLinks.nodeName = "Link" then
										sIcon = ""
										sText = ""
										sUrl = ""
										sHint = ""
										For Each xmlFichaLink In xmlFichaLinks.childNodes
											if xmlFichaLink.nodeName = "Text" then
												sText = xmlFichaLink.text
											elseif xmlFichaLink.nodeName = "Url" then
												sUrl = xmlFichaLink.text
											elseif xmlFichaLink.nodeName = "Icon" then
												sIcon = xmlFichaLink.text
											elseif xmlFichaLink.nodeName = "MouseOver" then
												sHint = xmlFichaLink.text
											end if
										Next

										if (sUrl <> "") then
											sData = "<a href='" & sUrl & "' class='link_classic2' title='" & sHint & "' target='_blank'>" & sText & "</a>"

											if (sIcon <> "") then
												sData = "<img src='" & sIcon & "' alt=''/>&nbsp;" & sData
											end if
											
											ficha = ficha & "<tr>"
											ficha = ficha & "<td class='td_ficha_esq esquerda'>Link</td>"
											ficha = ficha & "<td class='td_ficha_dir esquerda'>"&sData&"</td></tr>"
										end if
									end if
								Next
							end if
						Next
						
						Response.Write "<br /><table style='border-spacing: 2px; padding: 1px; width: 98%; display: inline-table;' class='td_grid_ficha_background'>"&ficha&"</table><br />"
						Response.Write "</td>"
						
						Response.Write "<td class='td_center_middle td_grid_ficha_background td_grid_ficha_borda' style='width: 104px'>"
						Response.Write "<table style='border-spacing: 1px; padding: 0px; background-color: #ffffff; display: inline-block;'>"

						Response.Write links

						Response.Write "</table>"
						Response.Write "</td>"		
						
						contRegistro = contRegistro + 1
					end if
				Next
			end if
		Next
	End if	
	Set xmlDoc = nothing
	Set xmlRoot = nothing		
End if

Response.Write "</table>"
%>
