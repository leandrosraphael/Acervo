<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top">
	<div id="loading-favorito" class="loading-favorito"><span class="span_imagem icon_16 mozilla_blu"></span></div>
	
<%
iSrvCorrente = cInt(iIndexSrv)

'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

if (listaSelecionada > 0)then
	
	Response.Write "<table class='max_width' style='border-spacing: 0'>"
	Response.Write "<tr>"
	Response.Write "	<td class='td_center_top'>"
	
	'******************************************************************************************
	Set ROService = ROServer.CreateService("Web_Consulta")
	sXMLFichas = ROService.MontaFichaFavoritos(listaSelecionada, global_idioma)
	Set ROService = nothing
	'******************************************************************************************
	
	if (repositorio_institucional = 0) then
		SET ROService = ROServer.CreateService("Web_Consulta")
		Set objCfgSolic = ROServer.CreateComplexType("TSolicConsulta")
		Set objCfgSolic = ROService.SolicConsulta
		iSolicConsulta = objCfgSolic.Solic_Consulta
		iTotalSolic = objCfgSolic.Solic_Consulta_Itens
		Set objCfgSolic = nothing
		Set ROService = nothing
	else
		iSolicConsulta = 0
		iTotalSolic = 0
	end if

	if iSolicConsulta = 1 then
		'Quando o terminal web estiver sendo executado fora da intranet, 
		'liberar o recurso de "Solicitar consulta".
		'IBCCRIM
		if (global_numero_serie = 4365) then
			iSolicConsulta = 1
		elseif (global_IP_Local = 0) then
			iSolicConsulta = 0
		end if
	end if
	
	'******************************************************************************************
	
	if len(trim(sXMLFichas)) <> 0 then
		if (left(sXMLFichas,5) = "<?xml") then
			Set xmlDoc = CreateObject("Microsoft.xmldom")
			xmlDoc.async = False
			xmlDoc.loadxml sXMLFichas
			Set xmlRoot = xmlDoc.documentElement
			if xmlRoot.nodeName = "Pagina" then
				quantidade = xmlRoot.attributes.getNamedItem("QUANTIDADE").value
				qtdLegislacao = xmlRoot.attributes.getNamedItem("QUANTIDADE_LEGISLACAO").value
			end if
			Set xmlDoc = nothing
			Set xmlRoot = nothing		
		end if
		Resultado = quantidade
		'******************************************************************************************
		%><!-- #include file="monta_vetores_favoritos.asp" --><%
		'******************************************************************************************

		Response.Write "<table class='tab_paginacao max_width remover_bordas_padding'>"
		Response.Write "<tr style='height: 26px'>"
		Response.Write "<td class='esquerda' style='width: 300px'>&nbsp;&nbsp;<span style='color: #666666'><b>"&getTermo(global_idioma, 8316, "Favoritos", 0)&":</b></span>&nbsp;&nbsp;"
		
		if quantidade = 1 then
			msg_sel = getTermo(global_idioma, 2945, "%s material selecionado", 0)
			msg_sel = Format(msg_sel, "<b>"&quantidade&"</b>")
			Response.Write msg_sel
		else
			msg_sel = getTermo(global_idioma, 2946, "%s materiais selecionados", 0)
			msg_sel = Format(msg_sel, "<b>"&quantidade&"</b>")
			Response.Write msg_sel
		end if		
			
		Response.Write "</td>"
		
		if ((iSolicConsulta = 1) and (global_hab_envio_email_ms = 1) and ((global_marc = 1) and (global_salvar_marc = 1) and (iNumOutros > 0)))then
			Response.Write "</tr>"
			Response.Write "<tr style='height: 26px'>"
		end if
		
		Response.Write "<td class='direita'>"
			
		if iSolicConsulta = 1 then
			if (quantidade - qtdLegislacao) <= 0 then
				sMsg = getTermo(global_idioma, 3970, "Não existem itens para serem solicitados", 0)
				if qtdLegislacao > 0 then
					sMsg = sMsg & " ("&getTermo(global_idioma, 3971, "As legislações não são solicitadas", 0)&")"
				end if
				sMsg = sMsg & "."
				Response.Write "<a class='link_serv' href=""javascript:alert('"&sMsg&"');""><span class='transparent-icon span_imagem icon_16 icon-small-document'></span>&nbsp;"&getTermo(global_idioma, 825, "Solicitar consulta", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			elseif (iTotalSolic > 0) and ((quantidade - qtdLegislacao) > iTotalSolic) then
				msg_solic = getTermo(global_idioma, 3972, "O número máximo de itens para a solicitação de consultas é %s.", 0)
				msg_solic = Format(msg_solic, iTotalSolic)
				sMsg = msg_solic & "\n" & getTermo(global_idioma, 3973, "Foram selecionados", 0) & " " & quantidade
				if qtdLegislacao > 0 then
					sMsg = sMsg & " ("&getTermo(global_idioma, 3971, "As legislações não são solicitadas", 0)&")"
				end if
				sMsg = sMsg & "."
				Response.Write "<a class='link_serv' href=""javascript:alert('"&sMsg&"');""><span class='transparent-icon span_imagem icon_16 icon-small-document'></span>&nbsp;"&getTermo(global_idioma, 825, "Solicitar consulta", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			else
				Response.Write "<a class='link_serv' href=""javascript:solic_sels("&listaSelecionada&", "&iSrvCorrente&");""><span class='transparent-icon span_imagem icon_16 icon-small-document'></span>&nbsp;"&getTermo(global_idioma, 825, "Solicitar consulta", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			end if
		end if
	
		if global_hab_envio_email_ms = 1 then
			Response.Write "<a class='link_serv' href=""javascript:envia_sels("&listaSelecionada&", "&iSrvCorrente&");""><span class='transparent-icon span_imagem icon_16 icon-small-message'></span>&nbsp;"&getTermo(global_idioma, 817, "Enviar por e-mail", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		end if

		Response.Write "<a class='link_serv' href=""javascript:imp_minhas_sels("&listaSelecionada&", "&iSrvCorrente&");""><span class='transparent-icon span_imagem icon_16 icon-small-print'></span>&nbsp;"&getTermo(global_idioma, 13, "Imprimir", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		
		if ((global_marc = 1) and (global_salvar_marc = 1) and (iNumOutros > 0)) then
			Response.Write "<a class='link_serv' href=""javascript:abreMarc(0,'sels');"" style='cursor:pointer'><span class='transparent-icon span_imagem icon_16 icon-small-save'></span>&nbsp;"&getTermo(global_idioma, 1334, "Salvar MARC", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		end if	

        Response.Write "<a class='link_serv' href=""javascript:limpar_lista_favorito("&listaSelecionada&", "&iSrvCorrente&");""><span class='transparent-icon span_imagem icon_16 icon-small-clear'></span>&nbsp;"&getTermo(global_idioma, 8349, "Limpar lista", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write "</tr></table>"
		
		sOrigem = "favoritos"			
		%><!-- #include file='grid_ficha.asp' --><%
		
		Response.Write "<table class='tab_paginacao max_width remover_bordas_padding'>"
		Response.Write "<tr style='height: 26px'>"
		Response.Write "<td class='direita'>"

		if iSolicConsulta = 1 then
			if (quantidade - qtdLegislacao) <= 0 then
				sMsg = getTermo(global_idioma, 3970, "Não existem itens para serem solicitados", 0)
				if qtdLegislacao > 0 then
					sMsg = sMsg & " ("&getTermo(global_idioma, 3971, "As legislações não são solicitadas", 0)&")"
				end if
				sMsg = sMsg & "."
				Response.Write "<a class='link_serv' href=""javascript:alert('"&sMsg&"');""><span class='transparent-icon span_imagem icon_16 icon-small-document'></span>&nbsp;"&getTermo(global_idioma, 825, "Solicitar consulta", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			elseif (iTotalSolic > 0) and ((quantidade - qtdLegislacao) > iTotalSolic) then
				msg_solic = getTermo(global_idioma, 3972, "O número máximo de itens para a solicitação de consultas é %s.", 0)
				msg_solic = Format(msg_solic, iTotalSolic)
				sMsg = msg_solic & "\n" & getTermo(global_idioma, 3973, "Foram selecionados", 0) & " " & quantidade
				if qtdLegislacao > 0 then
					sMsg = sMsg & " ("&getTermo(global_idioma, 3971, "As legislações não são solicitadas", 0)&")"
				end if
				sMsg = sMsg & "."
				Response.Write "<a class='link_serv' href=""javascript:alert('"&sMsg&"');""><span class='transparent-icon span_imagem icon_16 icon-small-document'></span>&nbsp;"&getTermo(global_idioma, 825, "Solicitar consulta", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			else
				Response.Write "<a class='link_serv' href=""javascript:solic_sels("&listaSelecionada&", "&iSrvCorrente&");""><span class='transparent-icon span_imagem icon_16 icon-small-document'></span>&nbsp;"&getTermo(global_idioma, 825, "Solicitar consulta", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			end if
		end if

		if global_hab_envio_email_ms = 1 then
			Response.Write "<a class='link_serv' href=""javascript:envia_sels("&listaSelecionada&", "&iSrvCorrente&");""><span class='transparent-icon span_imagem icon_16 icon-small-message'></span>&nbsp;"&getTermo(global_idioma, 817, "Enviar por e-mail", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		end if
		
		Response.Write "<a class='link_serv' href=""javascript:imp_minhas_sels("&listaSelecionada&", "&iSrvCorrente&");""><span class='transparent-icon span_imagem icon_16 icon-small-print'></span>&nbsp;"&getTermo(global_idioma, 13, "Imprimir", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	
		if ((global_marc = 1) and (global_salvar_marc = 1) and (iNumOutros > 0)) then
			Response.Write "<a class='link_serv' href=""javascript:abreMarc(0,'sels');"" style='cursor:pointer'><span class='transparent-icon span_imagem icon_16 icon-small-save'></span>&nbsp;"&getTermo(global_idioma, 1334, "Salvar MARC", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		end if
		
        Response.Write "<a class='link_serv' href=""javascript:limpar_lista_favorito("&listaSelecionada&", "&iSrvCorrente&");""><span class='transparent-icon span_imagem icon_16 icon-small-clear'></span>&nbsp;"&getTermo(global_idioma, 8349, "Limpar lista", 0)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write "</tr></table>"

		Response.Write "</tr>"
		Response.Write "</table>"
	else
		Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
		Response.Write "	<tr style='height: 26px'>"
		Response.Write "		<td>"		
		Response.Write "			<p class='centro'><b>" & getTermo(global_idioma, 1341, "Nenhum registro encontrado", 0) & "</b></p>"
		Response.Write "		</td>"
		Response.Write "	</tr>"
		Response.Write "	</table>"
					
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table>"
	end if
else
	Response.Write "<table class='tab_paginacao max_width remover_bordas_padding'>"
	Response.Write "<tr style='height: 26px'>"
	Response.Write "	<td>"
	Response.Write "		<p class='centro'><b>" & getTermo(global_idioma, 1341, "Nenhum registro encontrado", 0) & "</b></p>"
	Response.Write "	</td>"
	Response.Write "</tr>"
	Response.Write "</table>"		
end if

Response.Write "</div>"		
%>
</td>
</tr>
</table>
