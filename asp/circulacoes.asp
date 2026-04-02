<table class="tab_menu_servicos tab_emprestimo max_width max_height">
	<tbody>
		<tr>
			<td>
			<div class="div-servicos">
				<%
					acaoForm = Request.QueryString("acao")
					if (global_numero_serie = 6372) then
						if (acaoForm = "") then
							class_circulacao = "link_menu_serv2"
							class_emprestimo = "link_menu_serv1"
							class_devolucao = class_emprestimo
						else
							class_circulacao = "link_menu_serv1"
							if (acaoForm = "emprestimo") then
								class_emprestimo = "link_menu_serv2"
								class_devolucao = class_circulacao
							else
								class_devolucao = "link_menu_serv2"
								class_emprestimo = class_circulacao
							end if
						end if
				 %>
					<span style="width: 33.1%; float: left;">
						<a class="<%=class_circulacao%>" href="javascript:LinkCirculacoes(parent.hiddenFrame.modo_busca);">
							<span class="transparent-icon span_imagem icon_16 icon-small-circ"></span>&nbsp;Renovação
						</a>
					</span>

					<span style="width: 33.3%; float: left;">
						<a class="<%=class_devolucao%>" href="javascript:LinkDevolucao(parent.hiddenFrame.modo_busca);">
							<span class="transparent-icon span_imagem icon_16 icon-book-down"></span>&nbsp;Devolução
						</a>
					</span>

					<span style="width: 32.5%; float: left;">
						<a class="<%=class_emprestimo%>" href="javascript:LinkEmprestimo(parent.hiddenFrame.modo_busca);">
							<span class="transparent-icon span_imagem icon_16 icon-book-up"></span>&nbsp;Empréstimo
						</a>
					</span>
				<%end if %>
			</div>
			</td>
		</tr>
	</tbody>
</table>

<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top" style="display: block">
<script type="text/javascript">
var arPaginacao = new Array();
var iPagAtual   = 1;

function MarcarDesmarca() {
	for (var i = 0; i < document.frmCircula.elements.length; i++) {
		var x = document.frmCircula.elements[i];
		if (x.name.substring(0,2) == 'ck') {
			x.checked = document.frmCircula.selTudo.checked;
		}
	}
}

function LinkRenovar() {
	var iCount = 0;
	var sCods = "";
	for (var i = 0; i < document.frmCircula.elements.length; i++) {
		var x = document.frmCircula.elements[i];
		if (x.name.substring(0,2) == 'ck') {
			if (x.checked == true) {
				iCount++;
				if (sCods != "") {
					sCods = sCods + ",";
				}
				sCods = sCods + x.value;
			}
		}
	}
	if (iCount > 0) {
		if (parent.hiddenFrame.bloqueia_renovar == 0) {
			bloqueia_renovar(1);
			//Adequação Itaú
			//Filtra os serviços para exibir informações somente de uma biblioteca
			var iFiltroBib = 0;
			<% if ((global_numero_serie = 4090) OR (global_numero_serie = 4184)) then %>
				if (parent.hiddenFrame.iFixarBib == 1) {
					iFiltroBib = parent.hiddenFrame.geral_bib;
				}
			<% end if %>
			parent.mainFrame.location = 'index.asp?content=circulacoes&acao=renovacao&num_circulacao='+sCods+'&iBanner='+parent.hiddenFrame.iBanner + '&iIdioma='+parent.hiddenFrame.iIdioma+"&iFiltroBib="+iFiltroBib;
		} else {
			alert('<%=AddSlashes(getTermo(global_idioma, 1391, "A solicitação de renovação já foi feita e está em processamento. Por favor, aguarde.", 0))%>');
		}
	} else {
		alert('<%=AddSlashes(getTermo(global_idioma, 1390, "Nenhuma circulação foi selecionada para renovar.", 0))%>');
	}
}

function LinkCirculacao(iCirc,iPag,iTotalPag) {
	var dload = document.getElementById('circ_loading');
	if (dload != null) {
		dload.innerHTML = "<span class='span_imagem div_imagem_right icon_16 mozilla_blu '></span><span><%=getTermo(global_idioma, 32, "Carregando", 0)%>...</span>";
	}
	
	var sCod = arPaginacao[iPag-1];
	var url = 'asp/ajxCirculacoes.asp?circulacoes='+iCirc+'&pagina='+iPag+'&total='+iTotalPag+'&codigos='+sCod+'&iBanner='+parent.hiddenFrame.iBanner + '&iIdioma='+parent.hiddenFrame.iIdioma;
	var frm = document.getElementById('ajxCirc');
	if (frm != null) {
		frm.src   = url;
		iPagAtual = iPag;
	}
}

function LinkCircAnt(iCirc,iPag,iTotalPag) {
	if (iPag > 1) {
		LinkCirculacao(iCirc,iPag-1,iTotalPag);
	}
}

function LinkCircProx(iCirc,iPag,iTotalPag) {
	if (iPag < iTotalPag) {
		LinkCirculacao(iCirc,iPag+1,iTotalPag);
	}
}

function exibir_link_fgv(link) {
	var url = ext + "/pop_exibe_link.asp?link=" + escape(link);
	abrePopup(url, getTermo(global_frame.iIdioma, 9074, "Empréstimo de titulo digital", 0), 500, 390, true, true);
}

</script>
<iframe name="ajxCirc" id="ajxCirc" src="about:blank" style="display: none"></iframe>
<!-- DIV QUE ARMAZENA OS DADOS DO RECIBO DE RENOVAÇÃO -->
<div id="div_recibo" style="display:none;"></div>
<%

if Session("codigo_usuario") = "" then
	Response.Redirect("asp/logout.asp?logout=sim&modo_busca=" & GetModo_Busca & "&iBanner="&global_tipo_banner&"&iIdioma=" & global_idioma)
end if

if Session("nome_usuario") <> "" then

	if config_multi_servbib = 1 then
		iIndexSrv = Session("Servidor_Logado")

		if iIndexSrv = "" then
			iIndexSrv = 1
		end if

		'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
		%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	end if
	
	if (Request.QueryString("iFiltroBib") <> "") then
		iFiltroBib = Request.QueryString("iFiltroBib")
	else
		iFiltroBib = 0
	end if

	If (acaoForm = "renovacao") then %>
		<!-- #include file="renovacao.asp" -->
<%  elseif (acaoForm = "listarDevolucao") then %>
		<!-- #include file="listarDevolucao.asp" -->
<%  elseif (acaoForm = "devolver") then %>
		<!-- #include file="devolucao.asp" -->
<%  elseif (acaoForm = "emprestimo") then %>
		<!-- #include file="emprestimo.asp" -->
<%  else
		codigo_usu = Session("codigo_usuario")
		Nome_Usuario = Session("nome_usuario")
			
		codigo = codigo_usu
		On Error Resume Next
		Set ROService = ROServer.CreateService("Web_Consulta")
		
		XMLCircula = ROService.MontaCirculacoes(CLng(codigo),true,iFiltroBib,global_idioma)
		TrataErros(1)
		
		if (left(XMLCircula,5) = "<?xml") then
			Set xmlDoc = CreateObject("Microsoft.xmldom")
			xmlDoc.async = False
			xmlDoc.loadxml XMLCircula
			Set xmlRoot = xmlDoc.documentElement
			if xmlRoot.nodeName = "CIRCULACOES" then
				For Each xmlCirc In xmlRoot.childNodes
					'*************************************
					'Circulações em Aberto
					'*************************************
					if xmlCirc.nodeName  = "CIRCULACOES_ABERTAS" then
						QtdeCirc = xmlCirc.attributes.getNamedItem("QUANTIDADE").value
						sGridCircula = ""
						if QtdeCirc > 0 then
							Response.write "<form name='frmCircula'>"
							Response.write "<br /><p class='centro'><b>"&getTermo(global_idioma, 1392, "Circulações abertas", 0)&"</b><br /><br />"
							sGridCircula = sGridCircula & "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
							Registro = 1
							iQtdRenova = 0
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
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 25px'>&nbsp;"&sequencial_desc&"&nbsp;</td>"
											end if
										end if
										'*************************************
										'RENOVAÇÃO
										'*************************************
										if xmlCampos.nodeName  = "RENOVACAO" then
											renova_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
											pode_renovar = xmlCampos.attributes.getNamedItem("RENOVA").value
											renova = xmlCampos.attributes.getNamedItem("VALOR").value
											'Monta Registro
											if pode_renovar = 1 then
												iQtdRenova = iQtdRenova + 1
												sCheck = "<input type='checkbox' name='ck" & iQtdRenova & "' value='" & codigo_atual & "'>"
												sRegistro = sRegistro & "<td class='centro "&td_class&"'>"&sCheck&"</td>"
												'Monta Cabeçalho
												if Registro = 1 then
													sCheck = "<input type='checkbox' name='selTudo' title='"&getTermo(global_idioma, 1393, "Marcar/desmarcar todas as circulações.", 0)&"' onclick='MarcarDesmarca();'>"
												end if
											else
												sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;"&renova&"&nbsp;</td>"
											end if
											if Registro = 1 then
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 30px'>&nbsp;"&sCheck&"&nbsp;</td>"
											end if
										end if
										'*************************************
										'TITULO
										'*************************************
										if xmlCampos.nodeName  = "TITULO" then
											titulo_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value  
											titulo = xmlCampos.attributes.getNamedItem("VALOR").value
											titulo = "<a href='javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1,"&codigo_obra&",1,""circulacao""," & tipo_obra & ");'>" & titulo & "</a>"
											
											if xmlCampos.attributes.getNamedItem("DIGITAL").value = "1" then
												linkURI = xmlCampos.attributes.getNamedItem("URL_EMPRESTIMO").value
												titulo = titulo & "</br><div style='padding-left: 3px'>"& getTermo(global_idioma, 9075, "Clique", 0) & " "
												titulo = titulo & "<a href='javascript:exibir_link_fgv(""" & linkURI & """);'>"&getTermo(global_idioma, 9076, "aqui", 0)&"</a>"
												titulo = titulo & " " & getTermo(global_idioma, 9077, "para visualizar o link do empréstimo digital", 0) & "</div>"
											end if
																					
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
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 50px'>&nbsp;"&tombo_desc&"&nbsp;</td>"
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
											   codigoBib = xmlCampos.attributes.getNamedItem("CODIGO_BIBLIOTECA").value
											   biblioteca = xmlCampos.attributes.getNamedItem("NOME_BIBLIOTECA").value
												if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB_ATUAL").value = 1) then
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
													descricaoBib = descricaoBib & "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(codigoBibAtual) & ",0,0);'>("&getTermo(global_idioma, 9574, "Emp.", 0)& " " & bibliotecaAtual&")</a>"
												else
													descricaoBib = descricaoBib & "("&getTermo(global_idioma, 9574, "Emp.", 0)& " " & bibliotecaAtual & ")"
												end if
											end if

											'Monta Registro
											sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&descricaoBib&"&nbsp;</span></td>"
											'Monta Cabeçalho
											if Registro = 1 then
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 120px'>&nbsp;"&biblioteca_desc&"&nbsp;</td>"
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
												sCabecalho = sCabecalho & "<td class='centro td_tabelas_titulo' style='width: 100px'>&nbsp;"&datasai_desc&"&nbsp;</td>"
											end if
										end if
										'*************************************
										'DATA PREVISTA
										'*************************************
										if xmlCampos.nodeName  = "DATA_PREVISTA" then
											dataprev_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
											dataprev = xmlCampos.attributes.getNamedItem("VALOR").value
											atrasado = xmlCampos.attributes.getNamedItem("ATRASO").value
											'Monta Registro
											if atrasado = 1 then
												sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: red'>"&dataprev&"&nbsp;</td>"
											else
												sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&dataprev&"&nbsp;</td>"
											end if
											'Monta Cabeçalho
											if Registro = 1 then
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 100px'>&nbsp;"&dataprev_desc&"&nbsp;</td>"
											end if
										end if
									Next
									
									if Registro = 1 then
										sGridCircula = sGridCircula & "<tr style='height: 20px'>"
										sGridCircula = sGridCircula & sCabecalho
										sGridCircula = sGridCircula &  "</tr>"
									end if
									sGridCircula = sGridCircula & "<tr style='height: 20px'>" & sRegistro & "</tr>"
								end if
								
								Registro = Registro + 1
							Next
							sGridCircula = sGridCircula & "</table>"
							if iQtdRenova > 0 then
								sMsgRenovar = "<span style='color: red'><b>"&getTermo(global_idioma, 1394, "ATENÇÃO", 1)&": </b></span>"&getTermo(global_idioma, 1395, "Para validar uma renovação, selecione o(s) item(s) e clique na opção ""Renovar itens selecionados"".", 0)
								Response.Write "<table class='tab_mensagem' style='display: inline-table; margin-bottom: 10px;'><tr><td class='esquerda'>"&sMsgRenovar&"</td></tr></table><br />"
								sLinkRenovar = "<span class='transparent-icon span_imagem icon_16 icon-small-circ '></span>&nbsp;<a class='link_serv' href='javascript:LinkRenovar();' title='"&getTermo(global_idioma, 1396, "Para efetuar a renovação, selecione uma ou mais circulações abaixo e clique aqui.", 0)&"'>"&getTermo(global_idioma, 1397, "Renovar itens selecionados", 0)&"</a>"
								Response.write "<table style='width: 100%'>"
								Response.write "<tr style='height: 25px'><td class='td_servicos_titulo td_left_middle'>&nbsp;"&sLinkRenovar&"&nbsp;</td></tr></table>"
							end if
							Response.Write sGridCircula
							if iQtdRenova > 0 then
								sLinkRenovar = "<span class='transparent-icon span_imagem icon_16 icon-small-circ '></span>&nbsp;<a class='link_serv' href='javascript:LinkRenovar();' title='"&getTermo(global_idioma, 1396, "Para efetuar a renovação, selecione uma ou mais circulações abaixo e clique aqui.", 0)&"'>"&getTermo(global_idioma, 1397, "Renovar itens selecionados", 0)&"</a>"
								Response.write "<table style='width: 100%'>"
								Response.write "<tr style='height: 25px'><td class='td_servicos_titulo td_left_middle'>&nbsp;"&sLinkRenovar&"&nbsp;</td></tr></table>"
							end if							
							Response.write "</form><br />"
						else
							Response.write "<br /><br /><b>"&getTermo(global_idioma, 1392, "Circulações abertas", 0)&"</b>"
							msg_circ = getTermo(global_idioma, 1398, "Não existem circulações abertas para %s no momento.", 0)
							msg_circ = Format(msg_circ, "<b>"&Formata_Nome(Session("nome_usuario"),"inteiro")&"</b>")
							Response.write "<br /><br /><br />"&msg_circ&"<br /><br /><br />"
						end if
					end if
					'*************************************
					'Circulações de chave em Aberto
					'*************************************
					if xmlCirc.nodeName  = "CIRCULACOES_ABERTAS_CH" then
						QtdeCirc = xmlCirc.attributes.getNamedItem("QUANTIDADE").value
						sGridCircula = ""
						if QtdeCirc > 0 then
							Response.write "<p class='centro'><b>"&getTermo(global_idioma, 1399, "Circulações de chave abertas", 0)&"</b></p><br /><br />"
							sGridCircula = sGridCircula & "<table class='tab_circulacoes max_width' style='border-spacing: 1'>"
							Registro = 1
							iQtdRenova = 0
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
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 25px'>&nbsp;"&sequencial_desc&"&nbsp;</td>"
											end if
										end if
										'*************************************
										'DESCRIÇÃO
										'*************************************
										if xmlCampos.nodeName  = "DESCRICAO" then
											titulo_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
											titulo = xmlCampos.attributes.getNamedItem("VALOR").value
											'Monta Registro
											sRegistro = sRegistro & "<td class='esquerda "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&titulo&"&nbsp;</td>"
											'Monta Cabeçalho
											if Registro = 1 then
												sCabecalho = sCabecalho & "<td class='centro td_tabelas_titulo'>&nbsp;"&titulo_desc&"&nbsp;</td>"
											end if
										end if
										'*************************************
										'CODIGO
										'*************************************
										if xmlCampos.nodeName  = "CODIGO_CH" then
											codigo_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
											codigo = xmlCampos.attributes.getNamedItem("VALOR").value
											'Monta Registro
											sRegistro = sRegistro & "<td class='esquerda "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&codigo&"&nbsp;</td>"
											'Monta Cabeçalho
											if Registro = 1 then
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 50px'>&nbsp;"&codigo_desc&"&nbsp;</td>"
											end if
										end if
										'*************************************
										'TIPO
										'*************************************
										if xmlCampos.nodeName  = "TIPO" then
											tipo_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
											tipo = xmlCampos.attributes.getNamedItem("VALOR").value
											'Monta Registro
											sRegistro = sRegistro & "<td class='esquerda "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&tipo&"&nbsp;</td>"
											'Monta Cabeçalho
											if Registro = 1 then
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 110px'>&nbsp;"&tipo_desc&"&nbsp;</td>"
											end if
										end if
										'*************************************
										'DATA SAIDA
										'*************************************
										if xmlCampos.nodeName  = "DATA_SAIDA" then
											datasai_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
											datasai = xmlCampos.attributes.getNamedItem("VALOR").value
											'Monta Registro
											sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&datasai&"&nbsp;</td>"
											'Monta Cabeçalho
											if Registro = 1 then
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 95px'>&nbsp;"&datasai_desc&"&nbsp;</td>"
											end if
										end if

										'*************************************
										'DATA PREVISTA
										'*************************************
										if xmlCampos.nodeName  = "DATA_PREV" then
											dataprev_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
											dataprev = xmlCampos.attributes.getNamedItem("VALOR").value
											atrasado = xmlCampos.attributes.getNamedItem("ATRASO").value
											'Monta Registro
											if atrasado = 1 then
												sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: red'>"&dataprev&"&nbsp;</td>"
											else
												sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&dataprev&"&nbsp;</td>"
											end if
											'Monta Cabeçalho
											if Registro = 1 then
												sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 95px'>&nbsp;"&dataprev_desc&"&nbsp;</td>"
											end if
										end if
									Next
									
									if Registro = 1 then
										sGridCircula = sGridCircula & "<tr style='height: 20px'>"
										sGridCircula = sGridCircula & sCabecalho
										sGridCircula = sGridCircula &  "</tr>"
									end if
									sGridCircula = sGridCircula & "<tr style='height: 20px'>" & sRegistro & "</tr>"
								end if
								
								Registro = Registro + 1
							Next
							sGridCircula = sGridCircula & "</table>"
							Response.Write sGridCircula	
							Response.write "<br />"
						end if
					end if
				Next
			End if	
			Set xmlDoc = nothing
			Set xmlRoot = nothing	
		End if
		
		'*************************************
		'Monta a Paginação
		'*************************************
		XMLPaginacao = ROService.MontaCircPaginas(CLng(codigo_usu),iFiltroBib)
		aScript      = ""
		
		if (left(XMLPaginacao,5) = "<?xml") then
			Set xmlPag = CreateObject("Microsoft.xmldom")
			xmlPag.async = False
			xmlPag.loadxml XMLPaginacao
			Set xmlPagRoot = xmlPag.documentElement
			if xmlPagRoot.nodeName = "PAGINACAO" then
				CPAG = 1
				QtdePag   = xmlPagRoot.attributes.getNamedItem("PAGINAS").value
				TotalCirc = xmlPagRoot.attributes.getNamedItem("QUANTIDADE").value
				For Each xmlPagina In xmlPagRoot.childNodes
					if CPAG = 1 and xmlPagina.nodeName = "PAGINA" then
						sCodigo = xmlPagina.attributes.getNamedItem("CODIGOS").value
						CPAG    = CPAG + 1
					end if
					aScript = aScript & "arPaginacao.push('" & xmlPagina.attributes.getNamedItem("CODIGOS").value & "');"
				Next
			end if
			Set xmlPagRoot = nothing
			Set xmlPag = nothing

			XMLCirc = ROService.MontaCirculacoesPag(CStr(sCodigo),global_idioma)
			
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
								Response.Write "<script>"
								Response.Write aScript
								Response.Write "</script>"
							
								Response.write "<div id='circ_loading'>&nbsp;</div>"
								Response.write "<div id='div_circ'>"
								Response.write "<p class='centro'><b>"&getTermo(global_idioma, 1331, "Histórico de circulações", 0)&"</b></p><br /><br />"
								
								'*************************************
								'Exibe paginação
								'*************************************
								if (QtdePag > 1) then
									sNavegador = "<b>" & TotalCirc & "</b> "&getTermo(global_idioma, 1332, "circulações encontradas", 2)&"&nbsp;&nbsp;-&nbsp;&nbsp;<b>" & QtdePag & "</b> " & getTermo(global_idioma, 236, "Páginas", 0) & "&nbsp;&nbsp;&nbsp;&nbsp;"
									if (QtdePag > 5) then
										sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1325, "Primeira", 0)&"' style='cursor:pointer;' onClick=LinkCirculacao("&TotalCirc&",1,"&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-first '></span></a>"					
										sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1326, "Anterior", 0)&"' style='cursor:pointer;' onClick=LinkCircAnt("&TotalCirc&",1,"&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-previous '></span></a>"					
									end if
									iIDPag = 1
									do while ((iIDPag <= 5) and (iIDPag <= CInt(QtdePag)))
										sLink = "<a class='link_pag' title='"&getTermo(global_idioma, 1323, "Página", 0)&" "&iIDPag&"' href=""javascript:LinkCirculacao("&TotalCirc&","&iIDPag&","&QtdePag&");"">"
										if (1 = iIDPag) then
											sLink = sLink & "<b><span class='paginacao_link_atual'>"
										end if
										sLink = sLink & iIDPag
										if (1 = iIDPag) then
											sLink = sLink & "</span></b>"
										end if
										sNavegador = sNavegador & sLink & "</a>&nbsp;"
										iIDPag = iIDPag + 1
									loop
									if (QtdePag > 5) then
										sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1327, "Próxima", 0)&"' style='cursor:pointer;' onClick=LinkCircProx("&TotalCirc&",1,"&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-next '></span></a>"					
										sNavegador = sNavegador & "<a title='"&getTermo(global_idioma, 1328, "Última", 0)&"' style='cursor:pointer;' onClick=LinkCirculacao("&TotalCirc&","&QtdePag&","&QtdePag&") ><span class='transparent-icon span_imagem icon_16 icon-small-last '></span></a>"	
									end if
									Response.write "<table style='width: 100%'>"
									Response.write "<tr style='height: 25px'><td class='td_servicos_titulo td_left_middle'>&nbsp;"&sNavegador&"&nbsp;</td></tr></table>"
								end if
								
								Response.write "<table class='tab_circulacoes max_width' style='border-spacing: 1px'>"
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
													sCabecalho = sCabecalho & "<td class='centro td_tabelas_titulo' style='width: 25px'>&nbsp;"&sequencial_desc&"&nbsp;</td>"
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
													sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 50px'>&nbsp;"&tombo_desc&"&nbsp;</td>"
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
												   codigoBib = xmlCampos.attributes.getNamedItem("CODIGO_BIBLIOTECA").value
												   biblioteca = xmlCampos.attributes.getNamedItem("NOME_BIBLIOTECA").value 
													if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB").value = "1") then
														descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(codigoBib) & ",0,0);'>"&biblioteca&"</a>"
													else
														descricaoBib = biblioteca
													end if
												else
												   codigoBibAtual = xmlCampos.attributes.getNamedItem("CODIGO_BIB_ATUAL").value
												   bibliotecaAtual = xmlCampos.attributes.getNamedItem("NOME_BIB_ATUAL").value

												   codigoBib = xmlCampos.attributes.getNamedItem("CODIGO_BIBLIOTECA").value
												   biblioteca = xmlCampos.attributes.getNamedItem("NOME_BIBLIOTECA").value 
													if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB").value = "1") then
														descricaoBib = "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(codigoBib) & ",0,0);'>"&biblioteca&"</a>"
													else
														descricaoBib = biblioteca
													end if
													descricaoBib = descricaoBib & "<br />"
													if (xmlCampos.attributes.getNamedItem("EXIBIR_LINK_BIB_ATUAL").value = "1") then
														descricaoBib = descricaoBib & "<a class='link_classic2' style='cursor:pointer' href='javascript:InformacaoBiblioteca(" & Trim(codigoBibAtual) & ",0,0);'>(" & getTermo(global_idioma, 4558, "Emp.", 0) & " " & bibliotecaAtual&")</a>"
													else
														descricaoBib = descricaoBib & "(" & getTermo(global_idioma, 4558, "Emp.", 0) & " " & bibliotecaAtual & ")"
													end if
												end if

												'Monta Registro
												sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&descricaoBib&"&nbsp;</td>"
												'Monta Cabeçalho
												if Registro = 1 then
													sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 120px'>&nbsp;"&biblioteca_desc&"&nbsp;</td>"
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
													sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 95px'>&nbsp;"&datasai_desc&"&nbsp;</td>"
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
													sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 95px'>&nbsp;"&dataprev_desc&"&nbsp;</td>"
												end if
											end if
											'*************************************
											'DATA DEVOLUÇÃO
											'*************************************
											if xmlCampos.nodeName  = "DATA_DEV" then
												datadev_desc = xmlCampos.attributes.getNamedItem("DESCRICAO").value
												datadev = xmlCampos.attributes.getNamedItem("VALOR").value
												'Monta Registro
												sRegistro = sRegistro & "<td class='centro "&td_class&"'>&nbsp;<span style='color: "&fontcolor&"'>"&datadev&"&nbsp;</td>"
												'Monta Cabeçalho
												if Registro = 1 then
													sCabecalho = sCabecalho & "<td class='td_tabelas_titulo centro' style='width: 95px'>&nbsp;"&datadev_desc&"&nbsp;</td>"
												end if
											end if
										Next
										
										if Registro = 1 then
											Response.write "<tr style='height: 20px'>"
											Response.write sCabecalho
											Response.Write "</tr>"
										end if
										Response.write "<tr style='height: 20px'>" & sRegistro & "</tr>"
									end if
									
									Registro = Registro + 1
								Next						
								Response.write "</table><br />"						
							else
								Response.write "<br /><b>"&getTermo(global_idioma, 1331, "Histórico de circulações", 0)&"</b>"
								msg_circ = getTermo(global_idioma, 1400, "Não existe histórico de circulações para %s no momento.", 0)
								msg_circ = Format(msg_circ, "<b>"&Formata_Nome(Session("nome_usuario"),"inteiro")&"</b>")
								Response.write "<br /><br /><br />"&msg_circ&"<br /><br /><br />"
							end if
						end if
					Next
				End if	
				Set xmlDoc = nothing
				Set xmlRoot = nothing	
			End if
		End if
						
		if (left(XMLCircula,5) = "<?xml") then
			Set xmlDoc = CreateObject("Microsoft.xmldom")
			xmlDoc.async = False
			xmlDoc.loadxml XMLCircula
			Set xmlRoot = xmlDoc.documentElement
			if xmlRoot.nodeName = "CIRCULACOES" then
				For Each xmlCirc In xmlRoot.childNodes
					'*************************************
					'Pendências
					'*************************************
					if xmlCirc.nodeName  = "PENDENCIAS" then
						For Each xmlReg In xmlCirc.childNodes
							'*************************************
							'MULTA
							'*************************************
							if xmlReg.nodeName  = "MULTA" then
								Valor_Multa = xmlReg.attributes.getNamedItem("VALOR").value
                                
                                if (global_numero_serie = 4547) then
                                    msg_multa = "Você possui pendências financeiras."
                                else
                                    msg_multa = getTermo(global_idioma, 1453, "Você possui %smulta%s com valor de %s.", 0)
								    msg_multa = Format(msg_multa, "<span style='color: red'>|</span>|<b>"&Valor_Multa&"</b>")
                                end if
								
								Response.write "<p class='centro'><span style='color: "&fontcolor&"'>"&msg_multa&"</p>"
							end if
							'*************************************
							'BLOQUEIO
							'*************************************
							if xmlReg.nodeName  = "BLOQUEIO" then
								Data_Bloqueio = xmlReg.attributes.getNamedItem("VALOR").value
								Response.write "<p class='centro'><span style='color: "&fontcolor&"'><b>"&getTermo(global_idioma, 151, "Data de bloqueio", 0)&":</b> "&Data_Bloqueio&"</p>"
							end if
						Next
					end if
				Next
			End if	
			Set xmlDoc = nothing
			Set xmlRoot = nothing	
		End if
			
		Set ROService = nothing
	end if
else%>
	erro
<%end if%>
</td>
</tr>
</table>