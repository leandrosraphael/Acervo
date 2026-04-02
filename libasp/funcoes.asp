<%
'---------------------------------------------------------------------------------------------
'---------------------------------- F U N C O E S.A S P --------------------------------------
'---------------------------------------------------------------------------------------------
'  	    ______	  ______    __   _    _       _____   ______   ______   ______  ______        
'      / __   /  / __   /  / /  / \  / \     |  _  |  \   __\  \  __ \  \  ___\ \_   _\       
'     / /__/ /  / /__/ /  / /  /   \/   \    | |_| |   \  \___  \ \ \ \  \ \__     \ \        
'    / ____ /  / __   /  / /  / /\    /\ \   |  _  |     __   \  \ \ \ \  \  _\     \ \       
'   / /       / /  \ \  / /  / /  \__/  \ \  | | | |     \ \_\ \  \ \_\ \  \ \       \ \      
'  /_/       /_/   /_/ /_/  /_/          \_\ |_| |_|      \_____\  \_____\  \_\       \_\     
'                                                                                             
' Contém as funções utilizadas nas páginas 
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'                               S O P H I A    A C E R V O                                    
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

'------------------------------------------------------------------------
'Monta o combo de OBEJETOS ----------------------------------------------
'------------------------------------------------------------------------
Function ComboObjetos(xml,codSel)
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml xml
	Set xmlRoot = xmlDoc.documentElement
	
	ComboObjetos = ""
	if xmlRoot.nodeName  = "MATERIAIS" then
		sDescObjeto = xmlRoot.attributes.getNamedItem("DESC").value
	
		ComboObjetos = "<tr id='combo_material'><td class='direita'>" & sDescObjeto & "</td><td class='esquerda'>"
		ComboObjetos = ComboObjetos & "<select name='objeto' id='objeto' onChange='AtualizarComboMaterialOrdenacao(this.value)'>"
		
		For Each xmlPNode In xmlRoot.childNodes
			if xmlPNode.nodeName = "MATERIAL" then

				sDescObjeto  = xmlPNode.attributes.getNamedItem("DESC").value
				iCodObjeto   = xmlPNode.attributes.getNamedItem("COD").value
				ComboObjetos = ComboObjetos & "<option value='" & CStr(iCodObjeto) & "'"
				if (CInt(iCodObjeto) = CInt(codSel)) then
					ComboObjetos = ComboObjetos & " selected"
				end if
				ComboObjetos = ComboObjetos & ">" & sDescObjeto & "</option>"
			end if
		Next
		
		ComboObjetos = ComboObjetos & "</select></td></tr>"       
	end if
	
	Set xmlRoot = nothing
	Set xmlDoc = nothing
End Function

'------------------------------------------------------------------------
'Monta o combo de CONTEXTOS ---------------------------------------------
'------------------------------------------------------------------------
Function ComboContextos(xml,codSel)
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml xml
	Set xmlRoot = xmlDoc.documentElement
	
	ComboContextos = ""
	if xmlRoot.nodeName  = "CONTEXTOS" then
		iExibe = xmlRoot.attributes.getNamedItem("EXIBE").value
		sDescricao = xmlRoot.attributes.getNamedItem("DESC").value
	
		if (iExibe = 1) then
			ComboContextos = "<tr><td class='direita'>"&sDescricao&"</td><td class='esquerda'>"
			ComboContextos = ComboContextos & "<select name='contexto' name='id'>"
			
			For Each xmlPNode In xmlRoot.childNodes
				if xmlPNode.nodeName = "CONTEXTO" then
	
					sDescContexto  = xmlPNode.attributes.getNamedItem("DESC").value
					iCodContexto   = xmlPNode.attributes.getNamedItem("COD").value
					ComboContextos = ComboContextos & "<option value='" & CStr(iCodContexto) & "'"
					if (CInt(iCodContexto) = CInt(codSel)) then
						ComboContextos = ComboContextos & " selected"
					end if
					ComboContextos = ComboContextos & ">" & sDescContexto & "</option>"
				end if
			Next
			
			ComboContextos = ComboContextos & "</select></td></tr>"
		end if
	end if
	
	Set xmlRoot = nothing
	Set xmlDoc = nothing
End Function

'------------------------------------------------------------------------
'Exibe os Resultados ----------------------------------------------------
'------------------------------------------------------------------------
function FormataResultados(xml,paramURL)
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml xml
	Set xmlRoot = xmlDoc.documentElement
	
	FormataResultados = ""
	if xmlRoot.nodeName  = "FICHAS" then
		iQtd = CInt(xmlRoot.attributes.getNamedItem("QTDE_TOTAL").value)
		iPaginaAtual = CInt(xmlRoot.attributes.getNamedItem("PAGINA").value)
		iExibeMiniatura = xmlRoot.attributes.getNamedItem("EXIBE_MINIATURA").value
		iRegistrosPorPagina = CInt(xmlRoot.attributes.getNamedItem("REGISTROS_PAGINA").value)

		ID = ((iPaginaAtual - 1) * iRegistrosPorPagina) + 1

		sPaginacao = ""

		if iQtd > iRegistrosPorPagina then
			fQtdPagTemp = iQtd / iRegistrosPorPagina
			iQtdPag = Int(fQtdPagTemp)
			if fQtdPagTemp <> iQtdPag then
				iQtdPag = iQtdPag + 1
			end if

			iPrimeiraPagina = iPaginaAtual - 2
			if iPrimeiraPagina < 1 then
				iPrimeiraPagina = 1
			end if

			iUltimaPagina = iPrimeiraPagina + 4
			if iUltimaPagina > iQtdPag then
				iUltimaPagina = iQtdPag
				
				do while (((iUltimaPagina - iPrimeiraPagina) < 4) and (iPrimeiraPagina > 1)) 
					iPrimeiraPagina = iPrimeiraPagina - 1	
				loop
			end if

			sPaginacao = "<table cellspacing='1' cellpading='0' style='width: 100%;'><tr><td style='text-align: center;'>"
			if iPaginaAtual > 1 then
				sPaginacao = sPaginacao & "<a class='link_menu' href=""javascript:carregaPagina('"&paramURL&"',1);"">&lt;&lt;</a>&nbsp;&nbsp;<a class='link_menu' href=""javascript:carregaPagina('"&paramURL&"'," & (iPaginaAtual - 1) & ");"">&lt;</a>&nbsp;&nbsp;"
			end if
			for iNumPag = iPrimeiraPagina to iUltimaPagina
				if iNumPag = iPaginaAtual then
					sPaginacao = sPaginacao & "<b><u>" & iNumPag & "</u></b>&nbsp;&nbsp;"
				else
					sPaginacao = sPaginacao & "<a class='link_menu' href=""javascript:carregaPagina('"&paramURL&"',"&iNumPag&");"">" & iNumPag & "</a>&nbsp;&nbsp;"
				end if
			next
			if iPaginaAtual < iQtdPag then
				sPaginacao = sPaginacao & "<a class='link_menu' href=""javascript:carregaPagina('"&paramURL&"'," & (iPaginaAtual + 1) & ");"">&gt;</a>&nbsp;&nbsp;<a class='link_menu' href=""javascript:carregaPagina('"&paramURL&"'," & iQtdPag & ");"">&gt;&gt;</a>"
			end if
			sPaginacao = sPaginacao & "</td></tr><tr><td>&nbsp;</td></tr></table>"

		end if

		FormataResultados = FormataResultados & sPaginacao

		For Each xmlFicha In xmlRoot.childNodes
			if xmlFicha.nodeName = "FICHA" then
				CodItem = xmlFicha.attributes.getNamedItem("ITEM").value
				CodImg  = xmlFicha.attributes.getNamedItem("IMAGEM").value
				
				lupa = "imagens/icon-small-search.png"
				imgPref = "imagens/img_pref.gif"
				link = "asp/detalhes.asp?content=detalhe&codigo="&CodItem&paramURL

				FormataResultados = FormataResultados & "<div style='margin-left: 15px; margin-right: 15px; margin-bottom: 15px;'>"
				FormataResultados = FormataResultados & "<table style='border: 0; padding: 0;'><tr>"
				FormataResultados = FormataResultados & "<td style='width: 40px; padding: 0; vertical-align: top;text-align: center;'>" & ID & "</td>"
					
				if (iExibeMiniatura = 1) then
					
                    FormataResultados = FormataResultados & "<td class='td_grid_miniatura'>"

					if (CodImg <> "") AND (CodImg <> "0") then
						url_imagem_pref = "asp/imagem.asp?item=" & CodItem & "&imagem=" & CodImg & "&zoom=0"
						FormataResultados = FormataResultados & "<img src='" & url_imagem_pref & "' alt='' style='width: 100px;' />"
					else
						FormataResultados = FormataResultados & "<img src='imagens/imagem_padrao.jpg' alt='' />"
					end if

					FormataResultados = FormataResultados & "</td><td style='width: 650px; padding: 0; vertical-align: top;'>"
				
                else

                    FormataResultados = FormataResultados & "<td style='width: 760px; padding: 0; vertical-align: top;'>"
                end if

				FormataResultados = FormataResultados & "<table class='grid' cellspacing='1' cellpading='0' style='width: 100%;'>"
				
				For Each xmlCampos In xmlFicha.childNodes
					if xmlCampos.nodeName = "CAMPO" then
						sDescCampo  = xmlCampos.attributes.getNamedItem("DESCRICAO").value
						sValorCampo = TrocaTagMarcador(xmlCampos.attributes.getNamedItem("VALOR").value)
					
						FormataResultados = FormataResultados & "<tr>"
						FormataResultados = FormataResultados & "<td class='td_campo' width='150px'>&nbsp;"&sDescCampo&"</td>"
						FormataResultados = FormataResultados & "<td class='td_valor'>&nbsp;" & sValorCampo & "</td>"
						FormataResultados = FormataResultados & "</tr>"
					end if
				Next
				
				FormataResultados = FormataResultados & "</table>"
				
				FormataResultados = FormataResultados & "<td class='esquerda' style='width: 130px; padding: 0px 0px 0px 10px; vertical-align: top;'><a title='Detalhes...' class=""link_valor"" href='#' onclick=""javascript:ajxStartLoad('"&link&"');""><img border=0 class='transparent-icon' src='"&lupa&"'> Detalhes</a></td>"

				FormataResultados = FormataResultados & "</table>"
				
				FormataResultados = FormataResultados & "</div>"
				ID = ID + 1
			end if
		Next

		FormataResultados = FormataResultados & sPaginacao

	end if
	
	Set xmlRoot = nothing
	Set xmlDoc = nothing
end function

Function TrataErros
	If err.number <> 0 then
		msg_erro = err.source&" - "&err.Description
		numero_erro = "<b>Erro "&err.number&"</b> - "

		sMsg = "<div align='center'><img src='imagens/erro.gif'>&nbsp;<font color=#000033>"&numero_erro&msg_erro&"</font></div>"
		TrataErros = sMsg
	else
		TrataErros = ""
	End if
End Function

Function LiberarSessoes(bLogout)
	If(bLogout = true)then
		Session("loginext") = ""
		Session("logado") = ""
		Session("nome") = ""		
		Session("codigo") = 0
		Session("Autenticado") = ""
	End if
End Function

'------------------------------------------------------------------------
'Exibe quadro com dica  -------------------------------------------------
'------------------------------------------------------------------------
Function QuadroDica(xml)
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml xml
	Set xmlRoot = xmlDoc.documentElement
	
	QuadroDica = ""
	if xmlRoot.nodeName  = "DICA" then
	
		For Each xmlPNode In xmlRoot.childNodes
			if xmlPNode.nodeName = "TITULO" then
				sTitDica  = xmlPNode.attributes.getNamedItem("VALOR").value
			elseif xmlPNode.nodeName = "DESCRICAO" then
				sDescDica = xmlPNode.attributes.getNamedItem("VALOR").value			
			end if
		Next	
		
		Response.Write "<table class='grid' width='90%' style='margin-top: 16px'>"
		Response.Write "<tr class='tr_grid_cabecalho' style='height: 23px'>"
		Response.Write "<td class='td_grid_cabecalho'><center><b>"&sTitDica&"</b></center>"
		Response.Write "</td></tr>"
		Response.Write "<tr>"		
		Response.Write "<td class='td_valor'>"&sDescDica&"</td>"
		Response.Write "</tr>"		
		Response.Write "</table>"		
	end if
	
	Set xmlRoot = nothing
	Set xmlDoc = nothing
End Function
	
function TrocaTagMarcador(sXML)
	sXML = replace(sXML, "#D#", "<span class='destaca_palavras'>")
	sXML = replace(sXML, "#/D#", "</span>")
	TrocaTagMarcador = sXML
end function

function RemoveTagMarcador(sXML)
	sXML = replace(sXML, "#D#", "")
	sXML = replace(sXML, "#/D#", "")
	RemoveTagMarcador = sXML
end function


%>