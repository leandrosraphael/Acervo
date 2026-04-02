<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top" style="display: block">
<%
if Session("codigo_usuario") = "" OR Session("nome_usuario") = "" then
	Response.Redirect("asp/logout.asp?logout=sim&modo_busca="& GetModo_Busca)
else
	if config_multi_servbib = 1 then
		iIndexSrv = Session("Servidor_Logado")

		if iIndexSrv = "" then
			iIndexSrv = 1
		end if

		'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
		%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	end if

	Reservas()
end if
%>
</td>
</tr>
</table>
<%
Sub Reservas()

	if (Request.QueryString("iFiltroBib") <> "") then
		iFiltroBib = Request.QueryString("iFiltroBib")
	else
		iFiltroBib = 0
	end if
	
	'Busca todas as reservas do usuário
	Set ROService = ROServer.CreateService("Web_Consulta")	
	sXML = ROService.HistoricoReservas(Session("codigo_usuario"), iFiltroBib, global_idioma)
	Set ROService = nothing
	TrataErros(1)

	'Monta grid de reservas
	if (left(sXML,5) = "<?xml") then

		'Processa o XML com as reservas		
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml sXML
		Set xmlRoot = xmlDoc.documentElement

		'Verifica se usuário possui reservas
		if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
		
			'Cria grid com reservas
			response.write "<br /><p class='centro'><b>"&getTermo(global_idioma, 347, "Reservas", 0)&"</b><br /><br />"
			Response.write "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
			Response.write "<tr style='height: 20px'>"
			Response.write "<td class='td_tabelas_titulo centro' style='width: 25px'>&nbsp;#&nbsp;</td>"
			Response.write "<td class='td_tabelas_titulo centro' style='width: 25px'>&nbsp;"&getTermo(global_idioma, 2703, "Fila", 0)&"&nbsp;</td>"
			Response.write "<td class='td_tabelas_titulo centro'>&nbsp;"&getTermo(global_idioma, 177, "Título", 0)&"</td>"
			Response.write "<td class='td_tabelas_titulo centro' style='width: 40px'>"&getTermo(global_idioma, 2704, "Cód.", 0)&"</td>"
			Response.write "<td class='td_tabelas_titulo centro' style='width: 75px'>&nbsp;"&getTermo(global_idioma, 2705, "Data res.", 0)&"</td>"
			Response.write "<td class='td_tabelas_titulo centro' style='width: 65px'>&nbsp;"&getTermo(global_idioma, 2731, "Vencimento", 0)&"&nbsp;</td>"
			Response.write "<td class='td_tabelas_titulo centro' style='width: 75px'>&nbsp;"&getTermo(global_idioma, 241, "Situação", 0)&"&nbsp;</td>"
			if global_versao = vSOPHIA then
				Response.write "<td class='td_tabelas_titulo centro' style='width: 45px'>&nbsp;</td>"
			end if
			Response.write "</tr>"

			'Carrega cada uma das reservas
			For Each xmlReserva In xmlRoot.childNodes
				if xmlReserva.nodeName  = "RESERVA" then
					For Each xmlDados In xmlReserva.childNodes
						'Digital
						if xmlDados.nodeName  = "DIGITAL" then
							sDigital = xmlDados.attributes.getNamedItem("VALOR").value					
						end if

						'Sequencial
						if xmlDados.nodeName  = "SEQUENCIAL" then
							sSeq = xmlDados.attributes.getNamedItem("VALOR").value					
						end if
						
						'Posicao
						if xmlDados.nodeName  = "POSICAO" then
							sPosicao = xmlDados.attributes.getNamedItem("VALOR").value					
						end if
					
						'Titulo
						if xmlDados.nodeName  = "TITULO" then
							sTitulo = xmlDados.attributes.getNamedItem("VALOR").value
						end if

						'Material
						if xmlDados.nodeName  = "MATERIAL" then
							sMaterial = xmlDados.attributes.getNamedItem("VALOR").value
						end if

						'Ano da reserva
						if xmlDados.nodeName  = "ANO_RESERVA" then
							sAno = xmlDados.attributes.getNamedItem("VALOR").value
						end if

						'Volume da reserva
						if xmlDados.nodeName  = "VOLUME_RESERVA" then
							sVol = xmlDados.attributes.getNamedItem("VALOR").value
						end if

						'Número da reserva
						if xmlDados.nodeName  = "NUMERO_RESERVA" then
							sNum = xmlDados.attributes.getNamedItem("VALOR").value
						end if
						
						'Suporte da reserva
						if xmlDados.nodeName  = "SUPORTE" then
							sSuporte = xmlDados.attributes.getNamedItem("VALOR").value
						end if

						'Biblioteca
						if xmlDados.nodeName  = "BIB_RESERVA" then
							sBib = xmlDados.attributes.getNamedItem("VALOR").value
                            codigoBib = xmlDados.attributes.getNamedItem("CODIGO_BIBLIOTECA_RESERVA").value
						end if
						
						'Codigo da reserva
						if xmlDados.nodeName  = "CODIGO" then
							sCodigo = xmlDados.attributes.getNamedItem("VALOR").value
						end if

						'Código da obra
						if xmlDados.nodeName  = "CODIGO_OBRA" then
							sCod_Obra = xmlDados.attributes.getNamedItem("VALOR").value
						end if
					
						'Tipo da obra
						if xmlDados.nodeName  = "TIPO_OBRA" then
							sTipo_Obra = xmlDados.attributes.getNamedItem("VALOR").value
						end if

						'Data da reserva
						if xmlDados.nodeName  = "DATA_RESERVA" then
							sData = xmlDados.attributes.getNamedItem("VALOR").value
						end if
					
						'Vencimento da reserva
						if xmlDados.nodeName  = "VCTO_RESERVA" then						
							sVcto = xmlDados.attributes.getNamedItem("VALOR").value
						end if
					
						'Situacao da reserva
						if xmlDados.nodeName  = "SIT_RESERVA" then						
							sSit = xmlDados.attributes.getNamedItem("VALOR").value
						end if	
					Next

					'Define estilo da linha
					if (sSeq mod 2) > 0 then 'Linha ímpar
						td_class = "td_tabelas_valor2"
						link_class = "link_serv"
						img_excluir = "icon-small-delete-w2"
					else 'Linha par
						td_class = "td_tabelas_valor1"
						link_class = "link_serv"								
						img_excluir = "icon-small-delete-b-h"
					end if
			
					'Formata os dados do título da reserva com
					'Material, Biblioteca, Ano, Volume, Numero(ou Edição), Suporte
			
					sAux  = "<span style='color:#999999;font-weight:600'>"&getTermo(global_idioma, 175, "Material", 0)&":</span>&nbsp;"&sMaterial
	   			    if len(trim(sAno)) > 0 then
						sAux = sAux & "<span style='color:#999999;font-weight:600'> "&getTermo(global_idioma, 67, "Ano", 0)&":</span>&nbsp;"&sAno
					end if
					if len(trim(sVol)) > 0 then
						sAux = sAux & "<span style='color:#999999;font-weight:600'> "&getTermo(global_idioma, 355, "Vol.", 0)&":</span>&nbsp;"&sVol
					end if
					if len(trim(sNum)) > 0 then
						if cInt(sTipo_Obra) = 0 then
							sAux = sAux & "<span style='color:#999999;font-weight:600'> "&getTermo(global_idioma, 2732, "Núm.", 0)&":</span>&nbsp;"&sNum
						else
							sAux = sAux & "<span style='color:#999999;font-weight:600'> "&getTermo(global_idioma, 598, "Ed.", 0)&":</span>&nbsp;"&sNum
						end if
					end if
					if len(trim(sSuporte)) > 0 then
						sAux = sAux & "<span style='color:#999999;font-weight:600'> "&getTermo(global_idioma, 2733, "Sup.", 0)&":</span>&nbsp;"&sSuporte
					end if
					if len(trim(sBib)) > 0 then
                        descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(codigoBib) & ",0,0);'>"&sBib&"</a>"
						sAux = sAux & "<span style='color:#999999;font-weight:600'> "&descricaoBib&"</span>"
					end if

					'Inclui a reserva no grid
					Response.write "<tr>"
					Response.write "<td class='centro "&td_class&"'>&nbsp;<span style='color:"&fontcolor&"'>"&sSeq&"&nbsp;</td>"					
					Response.write "<td class='centro "&td_class&"'>&nbsp;<span style='color:"&fontcolor&"'>"&sPosicao&"&nbsp;</td>"					
					Response.write "<td class='esquerda "&td_class&"'>&nbsp;<span style='color:"&fontcolor&"'><a class='" & link_class & "' title='"&getTermo(global_idioma, 2555, "Abrir detalhes", 0)&"...' href=""javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,0,"&sCod_Obra&",1,'reserva',"&sTipo_Obra&");"">"&Replace(Replace(sTitulo,"<","&lt;"),">","&gt;")&"</a><br />&nbsp;"&sAux&" </td>"															
					Response.write "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&sCodigo&"&nbsp;</td>"										
					Response.write "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&sData&"&nbsp;</td>"					
					Response.write "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&sVcto&"&nbsp;</td>"										
					Response.write "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&sSit&"&nbsp;</td>"					
					if global_versao = vSOPHIA then
                        response.write "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'><a class='link_excluir' href='javascript:ConfirmaExclusaoReserva("&sCodigo&","&sDigital&");'><span class='transparent-icon span_imagem icon_16 "& img_excluir &"'></span></a>&nbsp;</td>"
					end if
					Response.Write "</tr>" 
				end if
			Next
			Response.write "</table>"
		else
			response.write "<br /><b>"&getTermo(global_idioma, 347, "Reservas", 0)&"</b>"
			msg_res = getTermo(global_idioma, 2734, "Não existem reservas para %s no momento.", 0)
			msg_res = Format(msg_res, "<b>"&Formata_Nome(Session("nome_usuario"),"inteiro")&"</b>")
			response.write "<br /><br /><br /><br />"&msg_res
		end if			
	else	 
		response.write "<br /><b>"&getTermo(global_idioma, 347, "Reservas", 0)&"</b>"
		msg_res = getTermo(global_idioma, 2734, "Não existem reservas para %s no momento.", 0)
		msg_res = Format(msg_res, "<b>"&Formata_Nome(Session("nome_usuario"),"inteiro")&"</b>")
		response.write "<br /><br /><br /><br />"&msg_res
	end if
End Sub
%>