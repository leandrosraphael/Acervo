<%
'------------------------------------------------------------------------
'Monta os campos de busca -----------------------------------------------
'------------------------------------------------------------------------
Function FormataCamposBusca(xml, Campo1, Campo2, Campo3, Campo4, Campo5, Campo6, Campo7, Campo8, iniId)
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml xml
	Set xmlRoot = xmlDoc.documentElement
                         
    Id = iniId
    FormataCamposBusca = ""
    
	if xmlRoot.nodeName  = "CAMPOS" then

		For Each xmlPNode In xmlRoot.childNodes
			if xmlPNode.nodeName = "CAMPO" then
				sTitCampo  = xmlPNode.attributes.getNamedItem("DESC").value
				sTipoCampo = xmlPNode.attributes.getNamedItem("TIPO").value
				
                if (iniId = "1") then
                   FormataCamposBusca = FormataCamposBusca & "<tr>"
                else
                   FormataCamposBusca = FormataCamposBusca & "<tr class='campo_material'>"
                end if
				FormataCamposBusca = FormataCamposBusca & "<td class='direita'>" & sTitCampo & "</td>"
				FormataCamposBusca = FormataCamposBusca & "<td class='esquerda'>"

				if ((Id = 1) and (campo1 <> "")) then
					sValorCampo = campo1
				elseif ((Id = 2) and (campo2 <> "")) then
					sValorCampo = campo2
				elseif ((Id = 3) and (campo3 <> "")) then
					sValorCampo = campo3
				elseif ((Id = 4) and (campo4 <> "")) then
					sValorCampo = campo4
				elseif ((Id = 5) and (campo5 <> "")) then
					sValorCampo = campo5
				elseif ((Id = 6) and (campo6 <> "")) then
					sValorCampo = campo6
				elseif ((Id = 7) and (campo7 <> "")) then
					sValorCampo = campo7
				elseif ((Id = 8) and (campo8 <> "")) then
					sValorCampo = campo8
				else
					sValorCampo = ""
				end if
			
				'-------------------------------------
				'CAMPO ALFANUMERICO / MEMO 
				'-------------------------------------
				if (CStr(sTipoCampo) = "1") OR (CStr(sTipoCampo) = "4") then		
					FormataCamposBusca = FormataCamposBusca & "<input type=""text"" name=""campo" & Id & _ 
										 """ id=""campo" & Id & """ style=""width: 285px"" value=""" & sValorCampo & """>"
					FormataCamposBusca = FormataCamposBusca & "<input type=""hidden"" name=""tipo_campo" & Id & """ id=""tipo_campo" & Id & """ value=""ALFA"">"
				'-------------------------------------
				'CAMPO DATA
				'-------------------------------------
				elseif (CStr(sTipoCampo) = "2") then
					if (sValorCampo = "") then
						cbOpc = "0"
						cmp1_d = ""
						cmp1_m = ""
						cmp1_a = ""
						cmp2_d = ""
						cmp2_m = ""
						cmp2_a = ""
					else
						arrCampo = split(sValorCampo, "|")
						cbOpc = arrCampo(0)
						cmp1  = split(arrCampo(1), "/")
						cmp2  = split(arrCampo(2), "/")
						
						cmp1_d = cmp1(0)
						cmp1_m = cmp1(1)
						cmp1_a = cmp1(2)
						
						if (cbOpc = "3") then
							cmp2_d = cmp2(0)
							cmp2_m = cmp2(1)
							cmp2_a = cmp2(2)
						else
							cmp2_d = ""
							cmp2_m = ""
							cmp2_a = ""
						end if
					end if
				
					FormataCamposBusca = FormataCamposBusca & "<select name=""cb_campo" & Id & """ id=""cb_campo" & Id & """ style='width: 93px' onchange=""habilitaEntre('data'," & Id & ");"">"
					if (cbOpc = "0") then
						FormataCamposBusca = FormataCamposBusca & "<option value='0' selected>Igual a</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='0'>Igual a</option>"
					end if
					if (cbOpc = "1") then
						FormataCamposBusca = FormataCamposBusca & "<option value='1' selected>Menor que</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='1'>Menor que</option>"
					end if
					if (cbOpc = "2") then
						FormataCamposBusca = FormataCamposBusca & "<option value='2' selected>Maior que</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='2'>Maior que</option>"
					end if
					if (cbOpc = "3") then
						FormataCamposBusca = FormataCamposBusca & "<option value='3' selected>Entre</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='3'>Entre</option>"
					end if
					FormataCamposBusca = FormataCamposBusca & "</select>"
					
					FormataCamposBusca = FormataCamposBusca & "&nbsp;"
					
					FormataCamposBusca = FormataCamposBusca & "<input type=""text"" maxlength=""2"" name=""campo" & Id & "_ini_d"" id=""campo" & Id & "_ini_d"" value=""" & cmp1_d & """ style='margin-left: 2px; margin-right: 2px; width: 25px'>/"
					FormataCamposBusca = FormataCamposBusca & "<input type=""text"" maxlength=""2"" name=""campo" & Id & "_ini_m"" id=""campo" & Id & "_ini_m"" value=""" & cmp1_m & """ style='margin-left: 2px; margin-right: 2px; width: 25px'>/"
					FormataCamposBusca = FormataCamposBusca & "<input type=""text"" maxlength=""4"" name=""campo" & Id & "_ini_a"" id=""campo" & Id & "_ini_a"" value=""" & cmp1_a & """ style='margin-left: 2px; width: 40px'>"
					
					if (cbOpc = "3") then
						FormataCamposBusca = FormataCamposBusca & "<span id='span_campo" & Id & "' style=""display: inline;"">"
					else
						FormataCamposBusca = FormataCamposBusca & "<span id='span_campo" & Id & "' style=""display: none;"">"
					end if
					FormataCamposBusca = FormataCamposBusca & "&nbsp;e&nbsp;"
					FormataCamposBusca = FormataCamposBusca & "<input type=""text"" maxlength=""2"" name=""campo" & Id & "_fim_d"" id=""campo" & Id & "_fim_d"" value=""" & cmp2_d & """ style='margin-left: 2px; margin-right: 2px; width: 25px'>/"
					FormataCamposBusca = FormataCamposBusca & "<input type=""text"" maxlength=""2"" name=""campo" & Id & "_fim_m"" id=""campo" & Id & "_fim_m"" value=""" & cmp2_m & """ style='margin-left: 2px; margin-right: 2px; width: 25px'>/"
					FormataCamposBusca = FormataCamposBusca & "<input type=""text"" maxlength=""4"" name=""campo" & Id & "_fim_a"" id=""campo" & Id & "_fim_a"" value=""" & cmp2_a & """ style='margin-left: 2px; width: 40px'>"
					FormataCamposBusca = FormataCamposBusca & "</span>"
					
					FormataCamposBusca = FormataCamposBusca & "<input type=""hidden"" name=""tipo_campo" & Id & """ id=""tipo_campo" & Id & """ value=""DATA"">"
				'-------------------------------------
				'CAMPO LOGICO
				'-------------------------------------
				elseif (CStr(sTipoCampo) = "3") then			
					FormataCamposBusca = FormataCamposBusca & "<select name=""campo" & Id & """ id=""campo" & Id & """ style='width: 289px'>"
					if (sValorCampo = "") then
						FormataCamposBusca = FormataCamposBusca & "<option value='' selected>Indiferente</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value=''>Indiferente</option>"
					end if
					if (sValorCampo = "0") then
						FormataCamposBusca = FormataCamposBusca & "<option value='0' selected>Não</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='0'>Não</option>"
					end if
					if (sValorCampo = "1") then
						FormataCamposBusca = FormataCamposBusca & "<option value='1' selected>Sim</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='1'>Sim</option>"
					end if
					FormataCamposBusca = FormataCamposBusca & "</select>"
					FormataCamposBusca = FormataCamposBusca & "<input type=""hidden"" name=""tipo_campo" & Id & """ id=""tipo_campo" & Id & """ value=""LOGICO"">"
				'-------------------------------------
				'CAMPO NUMERO
				'-------------------------------------
				elseif (CStr(sTipoCampo) = "5") then				
					if (sValorCampo = "") then
						cbOpc = "0"
						cmp1 = ""
						cmp2 = ""
					else
						arrCampo = split(sValorCampo, "|")
						cbOpc = arrCampo(0)
						cmp1  = arrCampo(1)
						if (cbOpc = "3") then
							cmp2 = arrCampo(2)
						else
							cmp2 = ""
						end if
					end if				
				
					FormataCamposBusca = FormataCamposBusca & "<select name=""cb_campo" & Id & """ id=""cb_campo" & Id & """ style='width: 93px' onchange=""habilitaEntre('num'," & Id & ");"">"
					if (cbOpc = "0") then
						FormataCamposBusca = FormataCamposBusca & "<option value='0' selected>Igual a</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='0'>Igual a</option>"
					end if
					if (cbOpc = "1") then
						FormataCamposBusca = FormataCamposBusca & "<option value='1' selected>Menor que</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='1'>Menor que</option>"
					end if
					if (cbOpc = "2") then
						FormataCamposBusca = FormataCamposBusca & "<option value='2' selected>Maior que</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='2'>Maior que</option>"
					end if
					if (cbOpc = "3") then
						FormataCamposBusca = FormataCamposBusca & "<option value='3' selected>Entre</option>"
					else
						FormataCamposBusca = FormataCamposBusca & "<option value='3'>Entre</option>"
					end if
					FormataCamposBusca = FormataCamposBusca & "</select>"
					
					FormataCamposBusca = FormataCamposBusca & "&nbsp;"					
					FormataCamposBusca = FormataCamposBusca & "<input type=""text"" name=""campo" & Id & "_ini"" id=""campo" & Id & "_ini"" size=""12"" value=""" & cmp1 & """>"
					
					if (cbOpc = "3") then
						FormataCamposBusca = FormataCamposBusca & "<span id='span_campo"  & Id & "' style=""display: inline;"">"
					else
						FormataCamposBusca = FormataCamposBusca & "<span id='span_campo" & Id & "' style=""display: none;"">"
					end if
					FormataCamposBusca = FormataCamposBusca & "&nbsp;e&nbsp;"
					FormataCamposBusca = FormataCamposBusca & "<input type=""text"" name=""campo" & Id & "_fim"" id=""campo" & Id & "_fim"" size=""12"" value=""" & cmp2 & """>"
					FormataCamposBusca = FormataCamposBusca & "</span>"
					
					FormataCamposBusca = FormataCamposBusca & "<input type=""hidden"" name=""tipo_campo" & Id & """ id=""tipo_campo" & Id & """ value=""NUM"">"
				
				'-------------------------------------
				'CAMPO TABELA
				'-------------------------------------
				elseif (CStr(sTipoCampo) = "6") then
					
					if (sValorCampo = "") then
						iValor = ""
					else
						arrCampo = split(sValorCampo, ",")
						iValor = arrCampo(1)
					end if										
					
					For Each xmlTabNode In xmlPNode.childNodes			

						if xmlTabNode.nodeName = "TABELA" then

							'TIPO DO CAMPO
							FormataCamposBusca = FormataCamposBusca & "<input type=""hidden"" name=""tipo_campo" & Id & _ 
												 """ id=""tipo_campo" & Id & """ value=""TABELA"">"
							'CÓDIGO DA TABELA
							FormataCamposBusca = FormataCamposBusca & "<input type=""hidden"" name=""cod_tabela" & Id & _
												 """ id=""cod_tabela" & Id & """ value=""" & xmlTabNode.attributes.getNamedItem("COD_TABELA").value & """>"
							'CAIXA DE SELEÇÃO
							FormataCamposBusca = FormataCamposBusca & "<select name=""cb_campo" & Id & _ 
												 """ id=""cb_campo" & Id & """ style='width: 289px'>"

							if (iValor = "") then
								FormataCamposBusca = FormataCamposBusca & "<option value='0' selected>Todos</option>"
							else
								FormataCamposBusca = FormataCamposBusca & "<option value='0'>Todos</option>"
							end if														

							'VALORES 
							For Each xmlValorNode In xmlTabNode.childNodes
								
								if xmlValorNode.nodeName = "VALOR" then
									
									if (iValor = xmlValorNode.attributes.getNamedItem("CODIGO").value) then 
										FormataCamposBusca = FormataCamposBusca & "<option value='" & xmlValorNode.attributes.getNamedItem("CODIGO").value & "' selected>" & _
															 xmlValorNode.attributes.getNamedItem("DESC").value & "</option>"
									else
										FormataCamposBusca = FormataCamposBusca & "<option value='" & xmlValorNode.attributes.getNamedItem("CODIGO").value & "' >" & _
															 xmlValorNode.attributes.getNamedItem("DESC").value & "</option>"
									end if

								end if
								
							Next
								
							FormataCamposBusca = FormataCamposBusca & "</select>"
						end if
					Next
				end if
				
				FormataCamposBusca = FormataCamposBusca & "</td></tr>"
				
				Id = Id + 1
			end if
		Next
        
	end if
	
	Set xmlRoot = nothing
	Set xmlDoc = nothing
End Function
%>