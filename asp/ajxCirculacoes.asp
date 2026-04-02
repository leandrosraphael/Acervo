<% Response.Buffer=True %>
<!DOCTYPE html>
<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
	<%
	sDiretorioArq = "asp"
	%>
	<!-- #include file="../config.asp" -->
	<!-- #include file="../idiomas/idiomas.asp" -->
	<!-- #include file="../libasp/header.asp" -->
	<!-- #include file="../libasp/funcoes.asp" -->
	<title>::<%=getTermo(global_idioma, 371, "Circulações", 0)%></title>
	<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css"> 
    <link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />
    <% if config_css_estatico = 1 then %>
		<link href="../inc/estilo.css" rel="stylesheet" type="text/css"> 
	<% else %>
		<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css">
	<% end if %>
	<link href="../inc/contraste.css" rel="stylesheet" type="text/css">
    <link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
	</head>
<body>
<%
if config_multi_servbib = 1 then
	iIndexSrv = Session("Servidor_Logado")

	if iIndexSrv = "" then
		iIndexSrv = 1
	end if

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
end if

On Error Resume Next

codigos = Request.QueryString("codigos")
QtdePag = Request.QueryString("total")
pag_atual = Request.QueryString("pagina")
TotalCirc = Request.QueryString("circulacoes")

Set ROService = ROServer.CreateService("Web_Consulta")
XMLCirc = ROService.MontaCirculacoesPag(CStr(codigos),global_idioma)

HTML_CIRC = ""

if (left(XMLCirc,5) = "<?xml") then
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml XMLCirc
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "CIRCULACOES" then
		For Each xmlCirc In xmlRoot.childNodes
			'*************************************
			'Histórico de Circulações
			'*************************************					
			if xmlCirc.nodeName  = "CIRCULACOES_HISTORICO" then
				QtdeCirc = xmlCirc.attributes.getNamedItem("QUANTIDADE").value
				
				if QtdeCirc > 0 then
					HTML_CIRC = HTML_CIRC & "<p class='centro'><b>"&getTermo(global_idioma,1331, "Histórico de circulações", 0)&"</b></p><br/><br/>"
					
					'*************************************
					'Exibe paginação
					'*************************************
					if (CInt(QtdePag) > 1) then
						sNavegador = "<b>" & TotalCirc & "</b> "&getTermo(global_idioma, 1332, "circulações encontradas", 2)&"&nbsp;&nbsp;-&nbsp;&nbsp;<b>" & QtdePag & "</b> "&getTermo(global_idioma, 236, "Páginas", 0)&"&nbsp;&nbsp;&nbsp;&nbsp;"
						if (CInt(QtdePag) > 5) then
							sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1325, "Primeira", 0)&"' style='cursor:pointer;' onClick=LinkCirculacao("&TotalCirc&",1,"&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-first '></span></a>"					
                            sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1326, "Anterior", 0)&"' style='cursor:pointer;' onClick=LinkCircAnt("&TotalCirc&","&pag_atual&","&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-previous '></span></a>"					
							if (CInt(pag_atual) <= 3) then
								iIDPag = 1
							elseif ((CInt(QtdePag) - CInt(pag_atual)) >= 2) then
								iIDPag = CInt(pag_atual) - 2
							else 
								iIDPag = CInt(QtdePag) - 4
							end if
						else
							iIDPag = 1
						end if
						do while ((iIDPag <= 5) and (iIDPag <= CInt(QtdePag)))
							sLink = "<a class='link_pag' title='"&getTermo(global_idioma, 1323, "Página", 0)&" "&iIDPag&"' href=""javascript:LinkCirculacao("&TotalCirc&","&iIDPag&","&QtdePag&");"">"
							if (CInt(pag_atual) = iIDPag) then
								sLink = sLink & "<b><span class='paginacao_link_atual'>"
							end if
							sLink = sLink & iIDPag
							if (CInt(pag_atual) = iIDPag) then
								sLink = sLink & "</span></b>"
							end if
							sNavegador = sNavegador & sLink & "</a>&nbsp;"
							iIDPag = iIDPag + 1
						loop
						if (QtdePag > 5) then
							sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1327, "Próxima", 0)&"' style='cursor:pointer;' onClick=LinkCircProx("&TotalCirc&","&pag_atual&","&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-next '></span></a>"					
                            sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1328, "Última", 0)&"' style='cursor:pointer;' onClick=LinkCirculacao("&TotalCirc&","&QtdePag&","&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-last '></span></a>"					
                        end if
						HTML_CIRC = HTML_CIRC & "<table style=' width: 100%;'>"
						HTML_CIRC = HTML_CIRC & "<tr style=' height: 25px;'><td class='td_servicos_titulo td_left_middle'>&nbsp;"&sNavegador&"&nbsp;</td></tr></table>"
					end if
					
					HTML_CIRC = HTML_CIRC & "<table class='tab_circulacoes max_width' style=' border-spacing: 1px;'>"
					Registro = 1						
					For Each xmlReg In xmlCirc.childNodes
						sCabecalho = ""
						sRegistro  = ""
						
						if (Registro mod 2) > 0 then '### IMPAR
							fontcolor = "#000000" 	
							td_class = "td_tabelas_valor2"
							link_class = "link_serv"
						else '### PAR
							fontcolor= "#000000" 
							td_class = "td_tabelas_valor1"
							link_class = "link_serv"								
						end if				
						
						if xmlReg.nodeName  = "CIRCULACAO" then
							For Each xmlCampos In xmlReg.childNodes
								'*************************************
								'INFORMAÇÕES DA OBRA
								'*************************************
								if xmlCampos.nodeName  = "INFO_OBRA" then
									codigo_obra = xmlCampos.attributes.getNamedItem("CODIGO").value
                                    tipo_obra = xmlCampos.attributes.getNamedItem("TIPO").value
                                end if
                                '*************************************
								'CODIGO
								'*************************************
								if xmlCampos.nodeName  = "CODIGO" then
									codigo_atual = xmlCampos.attributes.getNamedItem("VALOR").value
								end if
								'*************************************
								'SEQUENCIAL
								'*************************************
								if xmlCampos.nodeName  = "SEQUENCIAL" then
									sequencial_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
									sequencial = xmlCampos.attributes.getNamedItem("VALOR").value
									'Monta Registro
									sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&sequencial&"&nbsp;</td>"
									'Monta Cabeçalho
									if Registro = 1 then
										sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style=' width:25px;'>&nbsp;"&sequencial_desc&"&nbsp;</td>"
									end if
								end if
								'*************************************
								'TITULO
								'*************************************
								if xmlCampos.nodeName  = "TITULO" then
									titulo_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
									titulo = xmlCampos.attributes.getNamedItem("VALOR").value
									titulo = "<a href='javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1,"&codigo_obra&",1,""circulacao""," & tipo_obra & ");'>" & titulo & "</a>"
                                    'Monta Registro
									sRegistro = sRegistro & "<td class='esquerda "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&titulo&"&nbsp;</td>"
									'Monta Cabeçalho
									if Registro = 1 then
										sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro'>&nbsp;"&titulo_desc&"&nbsp;</td>"
									end if
								end if
                                '*************************************
     							'CLASSIFICAÇÃO  Não apresentar para Espanha e BN
	    						'*************************************
		    					if((global_espanha = 0) and (global_numero_serie <> 5592)) then
                                    if xmlCampos.nodeName  = "CLASSIFICACAO" then
				    				    classificacao_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
					    			    classificacao = xmlCampos.attributes.getNamedItem("VALOR").value
                                           'Monta Registro
							    	    sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&classificacao&"&nbsp;</td>"
									    'Monta Cabeçalho
									    if Registro = 1 then
								   		    sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 100px'>&nbsp;"&classificacao_desc&"&nbsp;</td>"
								        end if
								    end if
                                end if
								'*************************************
								'TOMBO OU CODIGO
								'*************************************
								if xmlCampos.nodeName  = "TOMBO" then
									tombo_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
									tombo = xmlCampos.attributes.getNamedItem("VALOR").value
									'Monta Registro
									sRegistro = sRegistro & "<td class='esquerda "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&tombo&"&nbsp;</td>"
									'Monta Cabeçalho
									if Registro = 1 then
										sCabecalho = sCabecalho & "<td class='td_tabelas_titulo esquerda' style='width: 50px'>&nbsp;"&tombo_desc&"&nbsp;</td>"
									end if
								end if
								'*************************************
								'CAMPO OPCIONAL
								'*************************************
								if xmlCampos.nodeName  = "CMP_OPCIONAL" then
									cmp_opc_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
									cmp_opc = xmlCampos.attributes.getNamedItem("VALOR").value
									'Monta Registro
									sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&cmp_opc&"&nbsp;</td>"
									'Monta Cabeçalho
									if Registro = 1 then
										sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 80px'>&nbsp;"&cmp_opc_desc&"&nbsp;</td>"
									end if
								end if
								'*************************************
								'BIBLIOTECA
								'*************************************
								if xmlCampos.nodeName  = "BIBLIOTECA" then
									biblioteca_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value

                                    if (xmlCampos.attributes.getNamedItem("CODIGO_BIB_ATUAL").value = xmlCampos.attributes.getNamedItem("CODIGO_BIBLIOTECA").value) then
                                       biblioteca = xmlCampos.attributes.getNamedItem("NOME_BIBLIOTECA").value 
                                       if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB").value = 1) then
                                          codigoBib = xmlCampos.attributes.getNamedItem("CODIGO_BIBLIOTECA").value
                                          descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(codigoBib) & ",0,0);'>"&biblioteca&"</a>"
                                       else
                                          descricaoBib = biblioteca
                                       end if
                                    else
                                       codigoBibAtual = xmlCampos.attributes.getNamedItem("CODIGO_BIB_ATUAL").value
                                       bibliotecaAtual = xmlCampos.attributes.getNamedItem("NOME_BIB_ATUAL").value

                                       codigoBib = xmlCampos.attributes.getNamedItem("CODIGO_BIBLIOTECA").value
                                       biblioteca = xmlCampos.attributes.getNamedItem("NOME_BIBLIOTECA").value 
                                       if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB").value = 1) then
                                            descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(codigoBib) & ",0,0);'>"&biblioteca&"</a>"
                                       else
                                            descricaoBib = biblioteca
                                       end if
                                       descricaoBib = descricaoBib & "<br />"
                                       if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB_ATUAL").value = 1) then
                                            descricaoBib = descricaoBib & "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(codigoBibAtual) & ",0,0);'>(" & getTermo(global_idioma, 4558, "Emp.", 0)& " " & bibliotecaAtual&")</a>"
                                       else
                                            descricaoBib = descricaoBib & "(" & getTermo(global_idioma, 4558, "Emp.", 0)& " " & bibliotecaAtual & ")"
                                       end if
                                    end if

									'Monta Registro
									sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&descricaoBib&"&nbsp;</td>"
									'Monta Cabeçalho
									if Registro = 1 then
										sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 120px;'>&nbsp;"&biblioteca_desc&"&nbsp;</td>"
									end if
								end if
								'*************************************
								'DATA DE SAIDA
								'*************************************
								if xmlCampos.nodeName  = "DATA_SAIDA" then
									datasai_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
									datasai = xmlCampos.attributes.getNamedItem("VALOR").value
									'Monta Registro
									sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&datasai&"&nbsp;</td>"
									'Monta Cabeçalho
									if Registro = 1 then
										sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 95px;'>&nbsp;"&datasai_desc&"&nbsp;</td>"
									end if
								end if
								'*************************************
								'DATA PREVISTA
								'*************************************
								if xmlCampos.nodeName  = "DATA_PREVISTA" then
									dataprev_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
									dataprev = xmlCampos.attributes.getNamedItem("VALOR").value
									'Monta Registro
									sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&dataprev&"&nbsp;</td>"
									'Monta Cabeçalho
									if Registro = 1 then
										sCabecalho = sCabecalho & "<td class='centro td_tabelas_titulo' style='width: 95px'>&nbsp;"&dataprev_desc&"&nbsp;</td>"
									end if
								end if
								'*************************************
								'DATA DEVOLUÇÃO
								'*************************************
								if xmlCampos.nodeName  = "DATA_DEV" then
									datadev_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
									datadev = xmlCampos.attributes.getNamedItem("VALOR").value
									'Monta Registro
									sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: '"&fontcolor&"'>"&datadev&"&nbsp;</td>"
									'Monta Cabeçalho
									if Registro = 1 then
										sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 95px'>&nbsp;"&datadev_desc&"&nbsp;</td>"
									end if
								end if
							Next
							
							if Registro = 1 then
								HTML_CIRC = HTML_CIRC & "<tr style=' height: 20px;'>"
								HTML_CIRC = HTML_CIRC & sCabecalho
								HTML_CIRC = HTML_CIRC & "</tr>"
							end if
							HTML_CIRC = HTML_CIRC & "<tr style=' height: 20px;'>" & sRegistro & "</tr>"
						end if
						
						Registro = Registro + 1
					Next						
					HTML_CIRC = HTML_CIRC & "</table><br />"						
				end if
			end if
		Next
	End if	
	Set xmlDoc = nothing
	Set xmlRoot = nothing	
End if

Set ROService = nothing
		
Response.Write "<script type='text/javascript'>"
Response.Write "	var HTML_CIRC = """ & replace(HTML_CIRC,"""","\""") & """;"
Response.Write "	var DIV = parent.document.getElementById('div_circ');"
Response.Write "	if (DIV != null) {"
Response.Write "		DIV.innerHTML = HTML_CIRC;"
Response.Write "	}"
Response.Write "	var DIV = parent.document.getElementById('circ_loading');"
Response.Write "	if (DIV != null) {"
Response.Write "		DIV.innerHTML = '&nbsp;';"
Response.Write "	}"
Response.Write "</script>"

%>
</body>
</html>