<!-- #include file="../libasp/funcoes.asp" -->
<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->
<%
if (sMsgErro <> "") then
	Response.Write sMsgErro
else%>
	<!-- #include file="../asp/ler_parametros_busca.asp" -->
	<%

    sURL_RESUMO = "&dados=" & Server.URLEncode(sDados) & "&objeto=" & iObjeto & _
				  "&contexto=" & iContexto & "&imagem=" & iImagem & _
				  "&campo1=" & Server.URLEncode(sCampo1) & "&campo2=" & Server.URLEncode(sCampo2) & _
				  "&campo3=" & Server.URLEncode(sCampo3) & "&campo4=" & Server.URLEncode(sCampo4) & _
				  "&campo5=" & Server.URLEncode(sCampo5) & "&campo6=" & Server.URLEncode(sCampo6) & _
				  "&campo7=" & Server.URLEncode(sCampo7) & "&campo8=" & Server.URLEncode(sCampo8) & _
				  "&campo_ordenacao=" & CodigoCampoOrdenacao & _
				  "&campo_ordem=" & OrdemPesquisa
	
    sURL_RESULTADOS = "&dados=" & Server.URLEncode(sDados) & "&objeto=" & iObjeto & _
					  "&contexto=" & iContexto & "&imagem=" & iImagem & _
				      "&campo1=" & Server.URLEncode(sCampo1) & "&campo2=" & Server.URLEncode(sCampo2) & _
				      "&campo3=" & Server.URLEncode(sCampo3) & "&campo4=" & Server.URLEncode(sCampo4) & _
				      "&campo5=" & Server.URLEncode(sCampo5) & "&campo6=" & Server.URLEncode(sCampo6) & _
				      "&campo7=" & Server.URLEncode(sCampo7) & "&campo8=" & Server.URLEncode(sCampo8) & _
					  "&tmp=" & sTabTmp & "&tmp_objeto=" & iTmpObjeto & "&tmp_contexto=" & iTmpContexto & "&pagina=" & iPagina & _
					  "&campo_ordenacao=" & CodigoCampoOrdenacao & _
					  "&campo_ordem=" & OrdemPesquisa

%>
<div id="divTituloPesquisa">
	<span style="float: left">
		<a class="link_menu" href="#" onClick=NovaPesquisa()>Home</a> > <a class="link_menu" href="#" onClick="LinkVoltar('detalhe_resultados','<%=sURL_RESUMO%>')">Resumo</a> > <a class="link_menu" href="#" onClick="LinkVoltar('resultados','<%=sURL_RESULTADOS%>')">Resultado</a> > Detalhes do item
	</span>
    <!-- #include file="../asp/botaoLogin.asp" -->
</div>
<br>
<%
	if IsNumeric(codItem) then
		'INÍCIO DO DESTAQUE
		Set ParamPesq = ROServer.CreateComplexType("TPesquisa")
		ParamPesq.sPalavraChave = sDados
		ParamPesq.iMaterial     = iObjeto
		ParamPesq.iContexto 	= iContexto
		ParamPesq.iImagem   	= iImagem
		ParamPesq.OPERADOR    	= Session("codigo") 
		ParamPesq.sCampo2		= sCampo2
		ParamPesq.sCampo3		= sCampo3
		ParamPesq.sCampo4		= sCampo4
		ParamPesq.sCampo5		= sCampo5
		ParamPesq.sCampo6		= sCampo6
		ParamPesq.sCampo7		= sCampo7
		ParamPesq.sCampo8		= sCampo8
		'FIM DO DESTAQUE

        codigoUsuario = Session("codigo")
		xmlDetalhe = ROService.GetDetalhe(CLng(codItem), ParamPesq, CLng(codigoUsuario))
		
		sMsgErro = TrataErros
		
		if (sMsgErro <> "") then
			Response.Write sMsgErro
		else	   
			Set xmlDoc = CreateObject("Microsoft.xmldom")
			xmlDoc.async = False
			xmlDoc.loadxml xmlDetalhe
			Set xmlRoot = xmlDoc.documentElement
			
			if xmlRoot.nodeName  = "ITEM" then
				iQtdImagens = 0
				redim arImagens(0)
				redim arImgContent(0)
				
				sAcervo  = xmlRoot.attributes.getNamedItem("ACERVO").value
				sColecao = xmlRoot.attributes.getNamedItem("COLECAO").value
		
				if (len(trim(sAcervo)) > 0) or (len(trim(sColecao)) > 0) then
					If len(trim(sAcervo)) > 0 then
						Response.Write "<b>Acervo: </b>"&sAcervo&"<br>"
					end if
					If len(trim(sColecao)) > 0 then
						Response.Write "<b>Coleção: </b>"&sColecao&"<br>"
					end if
					Response.Write "<br>"
				end if
		
				For Each xmlFichas In xmlRoot.childNodes
					'**********************************************
					' IMAGENS
					'**********************************************
					if xmlFichas.nodeName = "IMAGENS" then
						c = 0
						iQtdImagens = xmlFichas.attributes.getNamedItem("QTDE").value
						iZoomImg    = xmlFichas.attributes.getNamedItem("ZOOM").value
						redim arImagens(iQtdImagens-1)
						redim arImgContent(iQtdImagens-1)
						
						For Each xmlImagens In xmlFichas.childNodes
							if xmlImagens.nodeName = "IMAGEM" then
								arImagens(c)    = xmlImagens.attributes.getNamedItem("COD").value
								arImgContent(c) = xmlImagens.attributes.getNamedItem("ContentType").value
								c = c + 1
							end if
						Next
						
						if (iQtdImagens > 0) then
							Response.Write "<table>"
							Response.Write "<tr>"
							Response.Write "<td align=center>"
							Response.Write "<a href='#' class='link_menu' onClick=LinkAnt("&codItem&")>Anterior</a>&nbsp;"
							Response.Write "<img title='Ver ítem anterior...' src='imagens/icon-small-previous.png' class='transparent-icon' border=0 onClick='LinkAnt("&codItem&")'>"
							Response.Write "</td>"
							if (iZoomImg = 1) then
								Response.Write "<td align=center>&nbsp;&nbsp;"
								Response.Write "<a href='#' onclick='LinkImagem("&codItem&");'>"
								Response.Write "<img border=0 src='imagens/icon-small-search.png' class='transparent-icon' title='Ampliar Imagem...'>"
								Response.Write "</a>"
								Response.Write "&nbsp;&nbsp;</td>"
							end if
							Response.Write "<td align=center>"
							Response.Write "<img title='Ver próximo ítem...' src='imagens/icon-small-next.png' class='transparent-icon' border=0 onClick='LinkProx("&codItem&")'>"
							Response.Write "&nbsp;<a href='#' class='link_menu' onClick='LinkProx("&codItem&");'>Próximo</a>"
							Response.Write "</td>"
							Response.Write "</tr>"
							Response.Write "<tr>"
							Response.Write "<td colspan='3' height='23' align='center' id='iNumImg'>"
							Response.Write "<br><b>1</b>&nbsp;/&nbsp;<b>"&iQtdImagens&"</b>" 
							Response.Write "</td>"
							Response.Write "</tr>"
							Response.Write "</table>"
							
							Response.Write "<br><img id='img_item' src='asp/imagem.asp?item="&codItem&"&imagem="&arImagens(0)&"&zoom=0') width=290><br><br>"
						end if
					end if
					
					'**********************************************
					' MIDIAS
					'**********************************************
					if xmlFichas.nodeName = "MIDIAS" then
						Response.Write "<table class='grid' cellspacing=1 cellpading=0 width='95%'>"
						Response.Write "<tr class='tr_grid_cabecalho' height=23>"
						Response.Write "<td class='td_grid_cabecalho'>&nbsp;<b>Mídias</b>&nbsp;</td>"
						Response.Write "</tr>"
		
						For Each xmlMidias In xmlFichas.childNodes
							if xmlMidias.nodeName = "MIDIA" then
								iCodMidia  = xmlMidias.attributes.getNamedItem("COD").value
								sDescMidia = xmlMidias.attributes.getNamedItem("DESC").value
								sExtMidia  = xmlMidias.attributes.getNamedItem("EXT").value
								sTamMidia  = xmlMidias.attributes.getNamedItem("TAMANHO").value
                                usar_visualizador_pdf = xmlMidias.attributes.getNamedItem("USAR_VISUALIZADOR_PDF").value
							
								Response.Write "<tr>"
								Response.Write "<td class='td_valor'>"
								Response.Write "<img border=0 src='imagens/icon-small-openfolder.png' class='transparent-icon'>&nbsp;"
							
								if config_midia_offline = 1 then
										Response.Write "<a title='Clique aqui para visualizar a mídia' href='temp/"&iCodMidia&"."&sExtMidia&"' class='link_midias' Target='blank'>"
										Response.Write sDescMidia&"</a>"						
								else							
									if CStr(sTamMidia) <> "" then
										if CStr(sTamMidia) = "0" then
											Response.Write "<a title='Clique aqui para visualizar a mídia' href='#' class='link_midias' onClick=""javascript:alert('Mídia não disponível!');"">"
											Response.Write sDescMidia&"</a>"
											Response.Write " (Não disponível)"
										else
                                            if (usar_visualizador_pdf = 1) then
                                                Response.Write "<a title='Clique aqui para visualizar a mídia' href='#' class='link_midias' onClick=""LinkMidiaPdf("&iCodMidia&",'"&Server.URLEncode(sDescMidia)&"');"">" 
                                            else
											    Response.Write "<a title='Clique aqui para visualizar a mídia' href='#' class='link_midias' onClick=""LinkMidia("&iCodMidia&",'"&sExtMidia&"','"&Server.URLEncode(sDescMidia)&"');"">"
                                            end if
											Response.Write sDescMidia&"</a>"
											if CStr(sTamMidia) <> "-" then
												Response.Write " (" & sTamMidia & ")"
											end if										
										end if
									else
                                        if (usar_visualizador_pdf = 1) then
                                            Response.Write "<a title='Clique aqui para visualizar a mídia' href='#' class='link_midias' onClick=""LinkMidiaPdf("&iCodMidia&",'"&Server.URLEncode(sDescMidia)&"');"">" 
                                        else
										    Response.Write "<a title='Clique aqui para visualizar a mídia' href='#' class='link_midias' onClick=""LinkMidia("&iCodMidia&",'"&sExtMidia&"','"&Server.URLEncode(sDescMidia)&"');"">"
                                        end if
                                        Response.Write sDescMidia&"</a>"
									end if					
								end if
								
								Response.Write "</td>"
								Response.Write "</tr>"
							end if
						Next
						
						Response.Write "</table><br>"
					end if
					
					'**********************************************
					' LINKS
					'**********************************************
					if xmlFichas.nodeName = "LINKS" then
						Response.Write "<table class='grid' cellspacing=1 cellpading=0 width='95%'>"
						Response.Write "<tr class='tr_grid_cabecalho' height=23>"
						Response.Write "<td class='td_grid_cabecalho'>&nbsp;<b>Links</b>&nbsp;</td>"
						Response.Write "</tr>"
		
						For Each xmlLinks In xmlFichas.childNodes
							if xmlLinks.nodeName = "LINK" then
								iCodLink  = xmlLinks.attributes.getNamedItem("COD").value
								sDescLink = xmlLinks.attributes.getNamedItem("DESC").value
								sURLLink  = xmlLinks.attributes.getNamedItem("URL").value
							
								Response.Write "<tr>"
								Response.Write "<td class='td_valor'>"
								Response.Write "<img border=0 src='imagens/link.gif'>&nbsp;"
								Response.Write "<a title='Clique aqui para acessar o link' href='"&sURLLink&"' target='_blank' class='link_midias' href='#' onclick='ContarAcessoMidia(1140);'>"
								Response.Write sDescLink&"</a>"
								Response.Write "&nbsp;("&sURLLink&")"
								Response.Write "</td>"
								Response.Write "</tr>"
							end if
						Next
						
						Response.Write "</table><br>"
					end if
						
					'**********************************************
					' FICHAS
					'**********************************************
					if xmlFichas.nodeName = "FICHA" then	
						sDescFicha = xmlFichas.attributes.getNamedItem("DESC").value
						
						Response.Write "<table class='grid' cellspacing=1 cellpading=0 width='95%'>"
						Response.Write "<tr class='tr_grid_cabecalho' height=23>"
						Response.Write "<td class='td_grid_cabecalho' colspan=2>&nbsp;<b>"&sDescFicha&"</b>&nbsp;</td>"
						Response.Write "</tr>"
		
						For Each xmlCampos In xmlFichas.childNodes
							if xmlCampos.nodeName = "CAMPO" then
								sDescCampo  = xmlCampos.attributes.getNamedItem("DESC").value
								sValorCampo = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("VALOR").value)
								'sValorCampo = xmlCampos.attributes.getNamedItem("VALOR").value
								Response.Write "<tr>"
								Response.Write "<td class='td_campo' width='30%'>&nbsp;"&sDescCampo&"</td>"
								Response.Write "<td class='td_valor'>" & sValorCampo & "</td>"
								Response.Write "</tr>"
							end if
						Next
						
						Response.Write "</table><br>"
					end if
				Next
			end if
		end if
		
		Set ParamPesq = nothing
		
	end if
	
	Response.Write "<script type='text/javascript' id='script'>"
	Response.Write "	parent.hiddenFrame.img_total = " & iQtdImagens & ";"
	Response.Write "	parent.hiddenFrame.img_atual = 0;"
	
	if (iQtdImagens > 0) then
		aCodigos = ""
		aContent = ""
		for i = lbound(arImagens) to ubound(arImagens)
			if (aCodigos <> "") then
				aCodigos = aCodigos & ","
				aContent = aContent & ","
			end if
			aCodigos = aCodigos & CStr(arImagens(i))
			aContent = aContent & "'" & CStr(arImgContent(i)) & "'"
		next
		
		Response.Write "	parent.hiddenFrame.img_codigos = [" & aCodigos & "];"
		Response.Write "	parent.hiddenFrame.img_content = [" & aContent & "];"
	end if
	
	Response.Write "</script>"
	
end if

Set ROService = nothing
Set ROServer = nothing
%>