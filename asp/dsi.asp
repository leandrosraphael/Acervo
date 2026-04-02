<table class="max_width max_height">
<tr>
	<td class="td_padrao td_center_top" style="display: block">
	<%
	if Session("codigo_usuario") = "" then
   		Response.write "<br /><p class='centro'><b>"&getTermo(global_idioma, 1322, "Perfil de interesse", 0)&"</b><br /><br />"
		mensagens = getTermoHtml(global_idioma, 1461, "Para ter acesso aos serviços da biblioteca você deve estar logado no sistema.", 0)
		Response.Write "<table class='tab_mensagem'><tr><td class='esquerda'>"&mensagens&"</td></tr></table><br />"
	else 	
		codigo_usuario = Session("codigo_usuario")
	
		if config_multi_servbib = 1 then
			iIndexSrv = Session("Servidor_Logado")

			if iIndexSrv = "" then
				iIndexSrv = 1
			end if

			'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
			%><!-- #include file="../libasp/updChannelProperty.asp" --><%
		end if	

		On Error Resume Next
		Set ROService = ROServer.CreateService("Web_Consulta")
		XML_DSI = ROService.DadosDSI(codigo_usuario,global_idioma)
	
		TrataErros(1)
	
		Set ROService = nothing
	
		Tem_DSI   = false
		InfoEmail = ""
		InfoMat   = ""
		InfoBib   = ""
		InfoArea  = ""
		InfoAut   = ""
	
		if (left(XML_DSI,5) = "<?xml") then 
			Set xmlDoc = CreateObject("Microsoft.xmldom")
			xmlDoc.async = False
			xmlDoc.loadxml XML_DSI
			Set xmlRoot = xmlDoc.documentElement
			if xmlRoot.nodeName = "DSI" then
				For Each xmlDSI In xmlRoot.childNodes
					'*************************************************
					' E-MAIL 
					'*************************************************
					if xmlDSI.nodeName = "EMAIL" then
						desc_email = xmlDSI.attributes.getNamedItem("DESCRICAO").value
						val_email  = xmlDSI.attributes.getNamedItem("VALOR").value
				
						InfoEmail = "<tr style='height: 20px'>"
						InfoEmail = InfoEmail & "<td class='td_detalhe_descricao direita'>&nbsp;"&desc_email&"&nbsp;</td>"
						InfoEmail = InfoEmail & "<td class='td_detalhe_valor esquerda'>"
						InfoEmail = InfoEmail & "<table><tr><td style='width: 700px'>&nbsp;"
						InfoEmail = InfoEmail & val_email
						InfoEmail = InfoEmail & "</td><td style='width: 80px' class='direita'>"
                        if (global_atualizar_email = 1) then
						   InfoEmail = InfoEmail & "<a class='link_serv' href=""javascript:AlterarEmail(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-edit-w'></span>&nbsp;" & getTermo(global_idioma, 510, "Alterar", 0) & "</a>"
                        end if
						InfoEmail = InfoEmail & "</td></tr></table>"
						InfoEmail = InfoEmail & "</td>"
						InfoEmail = InfoEmail & "</tr>"
					end if
					'*************************************************
					' MATERIAIS
					'*************************************************
					if xmlDSI.nodeName = "MATERIAIS" then
						desc_mat = xmlDSI.attributes.getNamedItem("DESCRICAO").value
						Qtde     = xmlDSI.attributes.getNamedItem("QUANTIDADE").value
						Material = ""
						iSeq = 1
					
						if (Qtde <> "0") then
							For Each xmlMaterial In xmlDSI.childNodes
								if xmlMaterial.nodeName = "MATERIAL" then
									if Material <> "" then
										Material = Material & ", "
									end if
									Material = Material & iSeq & ".&nbsp;" & xmlMaterial.attributes.getNamedItem("VALOR").value
								
									iSeq = iSeq + 1
								end if
							Next
						else
							Material = "-"
						end if
					
						InfoMat = "<tr style='height: 20px'>"
						InfoMat = InfoMat & "<td class='td_detalhe_descricao direita'>&nbsp;"&desc_mat&"&nbsp;</td>"
						InfoMat = InfoMat & "<td class='td_detalhe_valor esquerda'>"
						InfoMat = InfoMat & "<table><tr><td style='width: 700px'>&nbsp;"
						InfoMat = InfoMat & Material
						InfoMat = InfoMat & "</td><td style='width: 80px' class='direita'>"
						InfoMat = InfoMat & "<a class='link_serv' href=""javascript:dsiSelMaterial(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-edit-w'></span>&nbsp;" & getTermo(global_idioma, 510, "Alterar", 0) & "</a>"
						InfoMat = InfoMat & "</td></tr></table>"
						InfoMat = InfoMat & "</td>"
						InfoMat = InfoMat & "</tr>"
					end if
					'*************************************************
					' BIBLIOTECAS
					'*************************************************
					if xmlDSI.nodeName = "BIBLIOTECAS" then
						desc_bib = xmlDSI.attributes.getNamedItem("DESCRICAO").value
						Qtde     = xmlDSI.attributes.getNamedItem("QUANTIDADE").value
						Biblioteca = ""
						iSeq = 1
					
						if (Qtde <> "0") then
							For Each xmlBib In xmlDSI.childNodes
								if xmlBib.nodeName = "BIBLIOTECA" then
									if Biblioteca <> "" then
										Biblioteca = Biblioteca & ", "
									end if
									Biblioteca = Biblioteca & iSeq & ".&nbsp;" & xmlBib.attributes.getNamedItem("VALOR").value
								
									iSeq = iSeq + 1
								end if
							Next
						else
							Biblioteca = "-"
						end if
					
						InfoBib = "<tr style='height: 20px'>"
						InfoBib = InfoBib & "<td class='td_detalhe_descricao direita'>&nbsp;"&desc_bib&"&nbsp;</td>"
						InfoBib = InfoBib & "<td class='td_detalhe_valor esquerda'>"
						InfoBib = InfoBib & "<table><tr><td style='width: 700px'>&nbsp;"
						InfoBib = InfoBib & Biblioteca
						InfoBib = InfoBib & "</td><td style='width: 80px' class='direita'>"
						InfoBib = InfoBib & "<a class='link_serv' href=""javascript:dsiSelBib(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-edit-w'></span>&nbsp;" & getTermo(global_idioma, 510, "Alterar", 0) & "</a>"
						InfoBib = InfoBib & "</td></tr></table>"
						InfoBib = InfoBib & "</td>"
						InfoBib = InfoBib & "</tr>"
					end if
					'*************************************************
					' AREAS
					'*************************************************
					if xmlDSI.nodeName = "AREAS" then
						desc_area = xmlDSI.attributes.getNamedItem("DESCRICAO").value
						Qtde      = xmlDSI.attributes.getNamedItem("QUANTIDADE").value
						Area = ""
						iSeq = 1
					
						if (Qtde <> "0") then
							For Each xmlArea In xmlDSI.childNodes
								if xmlArea.nodeName = "AREA" then
									if Area <> "" then
										Area = Area & ", "
									end if
									Area = Area & iSeq & ".&nbsp;" & xmlArea.attributes.getNamedItem("VALOR").value
								
									iSeq = iSeq + 1
								end if
							Next
						else
							Area = "-"
						end if
					
						InfoArea = "<tr style='height: 20px'>"
						InfoArea = InfoArea & "<td class='td_detalhe_descricao direita'>&nbsp;"&desc_area&"&nbsp;</td>"
						InfoArea = InfoArea & "<td class='td_detalhe_valor esquerda'>"
						InfoArea = InfoArea & "<table><tr><td style='width: 700px'>&nbsp;"
						InfoArea = InfoArea & Area
						InfoArea = InfoArea & "</td><td style='width: 80px' class='direita'>"
						InfoArea = InfoArea & "<a class='link_serv' href=""javascript:dsiSelArea(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-edit-w'></span>&nbsp;" & getTermo(global_idioma, 510, "Alterar", 0) & "</a>"
						InfoArea = InfoArea & "</td></tr></table>"
						InfoArea = InfoArea & "</td>"
						InfoArea = InfoArea & "</tr>"
					end if
					'*************************************************
					' AUTORIDADES
					'*************************************************
					if xmlDSI.nodeName = "AUTORIDADES" then
						desc_aut = xmlDSI.attributes.getNamedItem("DESCRICAO").value
					
						Autoridade = "<table class='max_width' style='border-spacing: 1px; padding:0'>"
						Autoridade = Autoridade & "<tr style='height: 20px'>"
						Autoridade = Autoridade & "<td class='td_tabelas_titulo centro'>"&desc_aut&"</td>"
						Autoridade = Autoridade & "<td class='td_tabelas_titulo centro' style='width: 90px'>&nbsp;"&getTermo(global_idioma, 83, "Tipo", 0)&"&nbsp;</td>"
						Autoridade = Autoridade & "<td class='td_tabelas_titulo centro' style='width: 35px'>&nbsp;</td>"
						Autoridade = Autoridade & "</tr>"
					
						iSeq = 1
					
						For Each xmlAut In xmlDSI.childNodes
							if (iSeq mod 2) > 0 then '### IMPAR
								td_class = "td_tabelas_valor2"
								link_class = "link_serv"
								img_excluir = "icon-small-delete-b"
							else '### PAR
								td_class = "td_tabelas_valor1"
								link_class = "link_serv"								
								img_excluir = "icon-small-delete-b-h"
							end if
					
							if xmlAut.nodeName = "AUTORIDADE" then
								Autoridade = Autoridade & "<tr style='height: 20px'>"
								Autoridade = Autoridade & "<td class='esquerda "&td_class&"'>"
								Autoridade = Autoridade &  "&nbsp;" & xmlAut.attributes.getNamedItem("VALOR").value
								Autoridade = Autoridade & "</td>"
							
								tipo_aut = xmlAut.attributes.getNamedItem("TIPO").value
								Select case tipo_aut
									case "100"
										desc_tipo = getTermo(global_idioma, 61, "Autor", 0)
									case "110"
										desc_tipo = getTermo(global_idioma, 62, "Instituição", 0)
									case "130"
										desc_tipo = getTermo(global_idioma, 70, "Título uniforme", 0)
									case "150"
										desc_tipo = getTermo(global_idioma, 72, "Assunto", 0)
								End Select
								Autoridade = Autoridade & "<td class='centro "&td_class&"'>"&desc_tipo&"</td>"
								Autoridade = Autoridade & "<td class='centro "&td_class&"'>&nbsp;"
								Autoridade = Autoridade & "<a class='" & link_class & "' href=""javascript:ExcluiDSI(parent.hiddenFrame.modo_busca,"&xmlAut.attributes.getNamedItem("CODIGO").value&");""><span class='transparent-icon span_imagem icon_16 " & img_excluir & " title='" & getTermo(global_idioma, 167, "Excluir", 0) & "'></span></a>"
								Autoridade = Autoridade & "&nbsp;</td>"
								Autoridade = Autoridade & "</tr>"
							
								iSeq = iSeq + 1
							end if
						Next
					
						Autoridade = Autoridade & "</table>"

						Tem_DSI = true
					end if
				Next
			end if
			Set xmlDoc = nothing
			Set xmlRoot = nothing
		end if
	
   		Response.write "<br /><p class='centro'><b>"&getTermo(global_idioma, 1322, "Perfil de interesse", 0)&"</b></p><br /><br />"

		Response.write "<table style='width: 100%; border-spacing: 1px; display: inline-table;'>"
	
		if (InfoEmail <> "") then
			Response.Write InfoEmail
		end if
	
		if (InfoMat <> "") then
			Response.Write InfoMat
		end if
	
		if (InfoBib <> "") then
			Response.Write InfoBib
		end if
	
		if (InfoArea <> "") then
			Response.Write InfoArea
		end if		
						
		Response.write "</table><br /><br />"
	
		if Tem_DSI = true then
			Response.write "<table style='width: 100%; border-spacing: 1px; display: inline-table;'>"
			Response.write "<tr style='height: 25px'>"
		
			Response.write "<td class='td_servicos_titulo centro'>"
			Response.write "<table class='max_width' style='border-spacing: 0'>"
			Response.write "<tr>"
			Response.Write "<td class='td_servicos_titulo esquerda' style='width: 300px'>&nbsp;&nbsp;"
			Response.Write "<a class='link_serv' href=""javascript:ExcluiDSI(parent.hiddenFrame.modo_busca,0);""><span class='transparent-icon span_imagem icon_16 icon-small-delete'></span>"
			Response.write "&nbsp;"&getTermo(global_idioma, 167, "Excluir", 0)&" "&getTermo(global_idioma, 1471, "interesse", 2)&"&nbsp;&nbsp;</a></td>"
			Response.Write "<td class='td_servicos_titulo direita'>"
			Response.write "<a class='link_serv' href=""javascript:LinkInteresse(parent.hiddenFrame.modo_busca+'&dsi=" & Codigo & "');""><span class='transparent-icon span_imagem icon_16 icon-small-edit'></span>"
			Response.write "&nbsp;"&getTermo(global_idioma, 510, "Alterar", 0)&" "&getTermo(global_idioma, 1471, "interesse", 2)&"&nbsp;&nbsp;</a></td>"
			Response.write "</tr></table>"
			Response.write "</td></tr>"
			Response.Write "<tr>"
			Response.Write "<td class='centro'><table class='tab_mensagem' style='width: 100% !important'><tr><td>" & Autoridade & "</td></tr></table></td>"
			Response.Write "</tr>"
			Response.write "</table><br />"
		else	
 			Response.write "<table style='width: 100%; border-spacing: 1px; display: inline-table;'>"
			Response.write "<tr style='height: 25px'>"
			Response.write "<td class='td_servicos_titulo direita'>"
			Response.write "<a class='link_serv' href=""javascript:LinkInteresse(parent.hiddenFrame.modo_busca);""><span class='transparent-icon span_imagem icon_16 icon-small-edit'></span>"
			Response.write "&nbsp;"&getTermo(global_idioma, 1472, "Novo", 0)&" "&getTermo(global_idioma, 1471, "interesse", 2)&"&nbsp;&nbsp;</a></td>"
			Response.write "</tr>"
			Response.write "</table>"
			Response.write "<table style='width: 100%; border-spacing: 1px; display: inline-table;'>"
			Response.write "<tr>"
			Response.write "<td class='centro td_tabelas_valor2' colspan=8>"
			msg_dsi = getTermo(global_idioma, 1473, "Não existe nenhum interesse cadastrado para %s no momento.", 0)
			msg_dsi = Format(msg_dsi, "<b>"&Formata_Nome(Session("nome_usuario"),"inteiro")&"</b>")
			Response.write msg_dsi
			Response.write "</td></tr>"
			Response.write "</table>"	
		end if
	end if

	%>

	</td>
</tr>
</table>