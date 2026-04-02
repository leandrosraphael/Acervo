<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
<%
sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%

iIndexSrv = IntQueryString("servidor", 1)
	
'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

codigo_obra 	= Request.querystring("codigo_obra")
tipo_obra 		= Request.querystring("tipo_obra")
codigo_usuario 	= Request.querystring("codigo_usuario")
modo_busca      = GetModo_Busca

%>
<script src="../scripts/jquery/jquery-3.2.1.min.js"></script>

<script type='text/javascript' src='../scripts/ajxDadosOpc.js'></script>
<script type='text/javascript'>

function atualiza_combos(cbAtual) {
	atualizaDadosOpc(<%=iIndexSrv%>, <%=codigo_usuario%>, <%=codigo_obra%>, cbAtual, "0", <%=tipo_obra%>);
}

function ValidaPreenchimento(desc_campo, campo){
	if ((campo != null) && (!campo.disabled) && (campo.value == "nenhum")){
		var termo = '<%=getTermo(global_idioma, 6660, "O campo %s é obrigatório.", 0)%>';
		termo = termo.replace(/%s/g, desc_campo);
		alert(termo);   
		campo.focus();  
		return false;
	}
	else{
		return true;
	}
}

function confirmaReserva(codigo_obra,codigo_usuario,bib_indica, tipo_obra) {

	ano    		= document.frm_dados_comp.ano.value;
	edicao 		= document.frm_dados_comp.edicao.value;
	volume 		= document.frm_dados_comp.volume.value;
	suporte		= document.frm_dados_comp.suporte.value;
	biblioteca 	= document.frm_dados_comp.biblioteca.value;
	edicao = RemoveAcentos(edicao);
	
	if (biblioteca=="nenhum" && bib_indica==1){
		alert('<%=getTermo(global_idioma, 1458, "O campo biblioteca é obrigatório.", 0)%>');
		document.frm_dados_comp.biblioteca.focus();
	} else {
		
		//UNICAMP
		if ((global_numero_serie == 4134) || (global_numero_serie == 5369)) {
			if (
				(!ValidaPreenchimento('<%=getTermo(global_idioma, 67, "Ano", 0)%>', document.frm_dados_comp.ano)) || 
				(!ValidaPreenchimento('<%=getTermo(global_idioma, 66, "Volume", 0)%>', document.frm_dados_comp.volume)) ||
				(!ValidaPreenchimento('<%=getTermo(global_idioma, 68, "Edição", 0)%>', document.frm_dados_comp.edicao)) ||
				(!ValidaPreenchimento('<%=getTermo(global_idioma, 69, "Suporte", 0)%>', document.frm_dados_comp.suporte))
			){
				return false;
			}
		}

		if ( ((global_numero_serie == 5859) || (global_numero_serie == 5895) || (global_numero_serie == 5896))){
			if 	(!ValidaPreenchimento('<%=getTermo(global_idioma, 66, "Volume", 0)%>', document.frm_dados_comp.volume)) {
				return false;
			}  

			if(tipo_obra == 0){
				if (
				(!ValidaPreenchimento('<%=getTermo(global_idioma, 67, "Ano", 0)%>', document.frm_dados_comp.ano)) || 
				(!ValidaPreenchimento('<%=getTermo(global_idioma, 101, "Número", 0)%>', document.frm_dados_comp.edicao))
				){
					return false;
				}
			}
		}
		
		url = "inclui_reserva.asp?codigo_usuario="+codigo_usuario+"&codigo_obra="+codigo_obra+"&biblioteca="+biblioteca+"&ano="+ano+"&volume="+volume+"&edicao="+edicao+"&suporte="+suporte+"&servidor=<%=iIndexSrv%>&modo_busca=<%=modo_busca%>&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>";
		document.location.href = url;
	}
}
</script>
<title>::<%=getTermo(global_idioma, 347, "Reservas", 0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />
<% if config_css_estatico = 1 then %>
	<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
<% else %>
	<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
<% end if %>
	<script type="text/javascript" src="../scripts/jquery/jquery-ui.min.js"></script>
	<script type="text/javascript" src="../scripts/jquery.multiselect.js"></script>
	<script type="text/javascript" src="../scripts/combo.init.js"></script>
	<link href="../scripts/css/jquery-ui-1.8.18.datepicker.css" rel="stylesheet" />

	<link href="../scripts/css/jquery.multiselect.css" rel="stylesheet" />
	<link href="../scripts/css/multiselect-custom.css" rel="stylesheet" />
	<link href="../scripts/css/pubtype-icons.css" rel="stylesheet" type="text/css" />
    <link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
    <link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
</head>
<body class="popup">
<script type="text/javascript">parent.fechaLoadingPopup();</script>
<br />
<div class="centro">
	<div style="height: 320px">
<% if tipo_obra = 1 then %>
	<b><%=getTermo(global_idioma, 64, "Dados opcionais", 0)%></b>
<% else %>
	<b><%=getTermo(global_idioma, 102, "Dados complementares", 0)%></b>
<% end if %>

<%
	'UNICAMP
	if ((global_numero_serie = 4134) or (global_numero_serie = 5369)) then
		Response.Write "<br /><br /><span style='font-size:7pt'>O preenchimento de todos os campos é obrigatório.</span>"
	end if
%>

<br />
<br />

<%
	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	
	xml_dados_opc = ROService.DadosOpcionais(codigo_usuario,codigo_obra,"","","",0,0)
	TrataErros(1)
		
	SET ROService = nothing

	num_anos = 0
	comboAnos = ""
	prim_ano = ""
	num_volumes = 0
	comboVolumes = ""
	prim_volume = ""
	num_numeros = 0
	comboNumeros = ""
	prim_numero = ""
	num_suporte = 0
	comboSuporte = ""
	prim_suporte = 0
	num_bib = 0
	comboBib = ""
	prim_bib = 0
	campoObrigatorio = ((global_numero_serie = 5859) or (global_numero_serie = 5895) or (global_numero_serie = 5896))
	
	if left(xml_dados_opc,5) = "<?xml" then
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml xml_dados_opc
		Set xmlRoot = xmlDoc.documentElement
		if xmlRoot.nodeName = "DADOS_OPC" then
			For Each xmlPNode In xmlRoot.childNodes
				'****************************************
				' MONTA O COMBO DE ANOS
				'****************************************
				if xmlPNode.nodeName = "ANOS" then
					comboAnos = "<select style='width:190px' id='ano' name='ano' class='styled_combo' onchange=""javascript:atualiza_combos('ano');"">"
					For Each xmlAnos In xmlPNode.childNodes
						'****************************************
						' ANOS 
						'****************************************
						if xmlAnos.nodeName = "ANO" then
							DescAno   = xmlAnos.attributes.getNamedItem("DESCRICAO").value
							if (DescAno = "") then
								' UNICAMP
								if ((global_numero_serie = 4134) or (global_numero_serie = 5369)) or (campoObrigatorio and tipo_obra = 0) then
									comboAnos = comboAnos & "<option value='nenhum' selected>"&getTermo(global_idioma, 128, "Selecionar", 0)&"</option>"
								else
									comboAnos = comboAnos & "<option value='nenhum' selected>"&getTermo(global_idioma, 8012, "Qualquer um", 0)&"</option>"
								end if
							else
								comboAnos = comboAnos & "<option value='"&DescAno&"'>"&DescAno&"</option>"
							end if
							num_anos  = num_anos + 1
							if (num_anos = 1) then
								prim_ano = DescAno
							end if
						end if
					Next
					comboAnos = comboAnos & "</select>"
				end if
				'****************************************
				' MONTA O COMBO DE VOLUMES
				'****************************************
				if xmlPNode.nodeName = "VOLUMES" then
					comboVolumes = "<select style='width:190px' id='volume' name='volume' class='styled_combo' onchange=""javascript:atualiza_combos('vol');"">"
					For Each xmlVols In xmlPNode.childNodes
						'****************************************
						' VOLUMES 
						'****************************************
						if xmlVols.nodeName = "VOLUME" then
							DescVol      = xmlVols.attributes.getNamedItem("DESCRICAO").value
							if (DescVol = "") then
								' UNICAMP e FGV 
								if ((global_numero_serie = 4134) or (global_numero_serie = 5369)) or (campoObrigatorio) then
									comboVolumes = comboVolumes & "<option value='nenhum' selected>"&getTermo(global_idioma, 128, "Selecionar", 0)&"</option>"
								else
									comboVolumes = comboVolumes & "<option value='nenhum' selected>"&getTermo(global_idioma, 8012, "Qualquer um", 0)&"</option>"
								end if
							else
								comboVolumes = comboVolumes & "<option value='"&DescVol&"'>"&DescVol&"</option>"
							end if
							num_volumes  = num_volumes + 1
							if (num_volumes = 1) then
								prim_volume = DescVol
							end if
						end if
					Next
					comboVolumes = comboVolumes & "</select>"
				end if
				'****************************************
				' MONTA O COMBO DE NÚMEROS / EDIÇÃO
				'****************************************
				if xmlPNode.nodeName = "NUMEROS" then
					comboNumeros = "<select style='width:190px' id='edicao' name='edicao' class='styled_combo' onchange=""javascript:atualiza_combos('num');"">"
					For Each xmlNums In xmlPNode.childNodes
						'****************************************
						' NUMEROS 
						'****************************************
						if xmlNums.nodeName = "NUMERO" then
							DescNum      = xmlNums.attributes.getNamedItem("DESCRICAO").value
							if (DescNum = "") then
								' UNICAMP
								if ((global_numero_serie = 4134) or (global_numero_serie = 5369)) or (campoObrigatorio and tipo_obra = 0) then
									comboNumeros = comboNumeros & "<option value='nenhum' selected>"&getTermo(global_idioma, 128, "Selecionar", 0)&"</option>"
								else
									comboNumeros = comboNumeros & "<option value='nenhum' selected>"&getTermo(global_idioma, 8012, "Qualquer um", 0)&"</option>"
								end if
							else
								comboNumeros = comboNumeros & "<option value='"&DescNum&"'>"&DescNum&"</option>"
							end if
							num_numeros  = num_numeros + 1
							if (num_numeros = 1) then
								prim_numero = DescNum
							end if
						end if
					Next
					comboNumeros = comboNumeros & "</select>"
				end if
				'****************************************
				' MONTA O COMBO DE SUPORTE
				'****************************************
				if xmlPNode.nodeName = "SUPORTES" then
					comboSuporte = "<select style='width:190px' id='suporte' name='suporte' class='styled_combo' onchange=""javascript:atualiza_combos('sup');"">"
					For Each xmlSup In xmlPNode.childNodes
						'****************************************
						' NUMEROS 
						'****************************************
						if xmlSup.nodeName = "SUPORTE" then
							CodSup       = xmlSup.attributes.getNamedItem("CODIGO").value
							DescSup      = xmlSup.attributes.getNamedItem("DESCRICAO").value
							if (CodSup = "") then
								' UNICAMP
								if ((global_numero_serie = 4134) or (global_numero_serie = 5369)) then
									comboSuporte = comboSuporte & "<option value='nenhum' selected>"&getTermo(global_idioma, 128, "Selecionar", 0)&"</option>"
								else
									comboSuporte = comboSuporte & "<option value='nenhum' selected>"&getTermo(global_idioma, 8012, "Qualquer um", 0)&"</option>"
								end if
							else
								comboSuporte = comboSuporte & "<option value='"&CodSup&"'>"&DescSup&"</option>"
							end if
							num_suporte  = num_suporte + 1
							if (num_suporte = 1) then
								prim_suporte = CodSup
							end if
						end if
					Next
					comboSuporte = comboSuporte & "</select>"
				end if
				'****************************************
				' MONTA O COMBO DE BIBLIOTECAS
				'****************************************
				if xmlPNode.nodeName = "BIBLIOTECAS" then
					comboBib = "<select style='width:190px' id='biblioteca' name='biblioteca' class='styled_combo' onchange=""javascript:atualiza_combos('bib');"">"
					For Each xmlBibs In xmlPNode.childNodes
						'****************************************
						' BIBILIOTECAS 
						'****************************************
						if xmlBibs.nodeName = "BIBLIOTECA" then
							DescBib  = xmlBibs.attributes.getNamedItem("DESCRICAO").value
							CodBib   = xmlBibs.attributes.getNamedItem("CODIGO").value
							if (CodBib = "") then
								' UNICAMP
								if ((global_numero_serie = 4134) or (global_numero_serie = 5369)) then
									comboBib = comboBib & "<option value='nenhum' selected>"&getTermo(global_idioma, 128, "Selecionar", 0)&"</option>"
								else
									comboBib = comboBib & "<option value='nenhum' selected>"&getTermo(global_idioma, 8012, "Qualquer um", 0)&"</option>"
								end if
							else
								comboBib = comboBib & "<option value='"&CodBib&"'>"&DescBib&"</option>"
							end if
							num_bib  = num_bib + 1
							if (num_bib = 1) then
								prim_bib = CodBib
							end if
						end if
					Next
					comboBib = comboBib & "</select>"
				end if
			Next
		end if
	end if
	
	If num_anos > 1 OR num_volumes > 1 OR num_numeros > 1 OR num_suporte > 1 OR num_bib > 1 then
		If IsNumeric(codigo_obra) AND IsNumeric(codigo_usuario) then
            Response.Write "<form name='frm_dados_comp' method='POST'>"
			Response.Write "<table style='padding: 2px; border-spacing: 2px; display: inline-table;' class='centro'>"
			
			If num_bib > 1 then
				Response.Write "<tr>"&vbcrlf
				Response.Write "<td style='padding: 2px'>"&getTermo(global_idioma, 3, "Biblioteca", 0)&"&nbsp;</td>"&vbcrlf
				Response.Write "<td style='padding: 2px'>"&vbcrlf
				Response.Write comboBib
				Response.Write "</td></tr>"&vbcrlf
			else
				if (num_bib = 1) then
					Response.Write "<input type='hidden' name='biblioteca' value='" & prim_bib & "'/>"
				else
					Response.Write "<input type='hidden' name='biblioteca' value=''/>"
				end if
			end if
			
			If (num_volumes > 1) then
				Response.Write "<tr>"&vbcrlf
				Response.Write "<td class='esquerda' style='padding: 2px'>"&getTermo(global_idioma, 66, "Volume", 0)&"&nbsp;</td>"&vbcrlf
				Response.Write "<td style='padding: 2px'>"&vbcrlf
				Response.Write comboVolumes
				Response.Write "</td></tr>"&vbcrlf
			else
				'Adequação CTA -  se possuir unico valor insere a reserva com o valor
				if (global_numero_serie = 3156) and (num_volumes = 1) and (prim_volume <> "") then
					Response.Write "<input type='hidden' name='volume' value='" & prim_volume & "'/>"
				else
					Response.Write "<input type='hidden' name='volume' value=''/>"
				end if
			end if
			If (num_anos > 1) then
				Response.Write "<tr>"&vbcrlf
				Response.Write "<td class='esquerda' style='padding: 2px'>"&getTermo(global_idioma, 67, "Ano", 0)&"&nbsp;</td>"&vbcrlf
				Response.Write "<td style='padding: 2px'>"&vbcrlf
				Response.Write comboAnos
				Response.Write "</td></tr>"&vbcrlf
			else
				Response.Write "<input type='hidden' name='ano' value=''>"
			end if
			If (num_numeros > 1) then
				Response.Write "<tr>"&vbcrlf
				if tipo_obra = 1 then
					Response.Write "<td class='esquerda' style='padding: 2px'>"&getTermo(global_idioma, 68, "Edição", 0)&"&nbsp;</td>"&vbcrlf
				else
					Response.Write "<td class='esquerda' style='padding: 2px'>"&getTermo(global_idioma, 101, "Número", 0)&"&nbsp;</td>"&vbcrlf
				end if
				Response.Write "<td style='padding: 2px'>"&vbcrlf
				Response.Write comboNumeros
				Response.Write "</td></tr>"&vbcrlf
			else
				'Adequação CTA -  se possuir unico valor insere a reserva com o valor
				if (global_numero_serie = 3156) and (num_numeros = 1) and (prim_numero <> "") then
					Response.Write "<input type='hidden' name='edicao' value='" & prim_numero & "'/>"
				else
					Response.Write "<input type='hidden' name='edicao' value=''/>"
				end if
			end if
			If (num_suporte > 1) then
				Response.Write "<tr>"&vbcrlf
				Response.Write "<td class='esquerda' style='padding: 2px'>"&getTermo(global_idioma, 69, "Suporte", 0)&"&nbsp;</td>"&vbcrlf
				Response.Write "<td style='padding: 2px'>"&vbcrlf
				Response.Write comboSuporte
				Response.Write "</td></tr>"&vbcrlf
			else
				'Adequação CTA -  se possuir unico valor insere a reserva com o valor
				if (global_numero_serie = 3156) and (num_suporte = 1) and (CStr(prim_suporte) <> "0") and (CStr(prim_suporte) <> "") then
					Response.Write "<input type='hidden' name='suporte' value='" & prim_suporte &"'>"
				else
					Response.Write "<input type='hidden' name='suporte' value=''>"
				end if
			end if
						
			Response.Write "</table></div>"
			Response.Write "<div class='centro'><br />"
			Response.Write "<input type='hidden' name='veio_de' value='form'>"
			Response.Write "<input type='button' name='submit' value='"&getTermo(global_idioma, 1459, "Reservar", 0)&"' onclick='javascript:confirmaReserva("&codigo_obra&","&codigo_usuario&","&global_bib_indica&"," &tipo_obra&" );'>"
			Response.Write "&nbsp;&nbsp;&nbsp;"
			Response.Write "<input type='button' name='cancelar' value='"&getTermo(global_idioma, 5, "Cancelar", 0)&"' onclick='javascript:parent.fechaPopup(""asp"");'>"
			Response.Write "</div>"
		end if	
	else
		url = "inclui_reserva.asp?codigo_obra="&codigo_obra&"&codigo_usuario="&codigo_usuario&"&modo_busca="&modo_busca&"&servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma
		if (num_bib = 1) then
			url = url & "&biblioteca=" & prim_bib
		end if
		'Adequação CTA -  se possuir unico valor insere a reserva com o valor
		if (global_numero_serie = 3156) and (num_volumes = 1) and (prim_volume <> "") then
			url = url & "&volume=" & prim_volume
		end if
		if (global_numero_serie = 3156) and (num_numeros = 1) and (prim_numero <> "") then
			url = url & "&edicao=" & prim_numero
		end if
		if (global_numero_serie = 3156) and (num_suporte = 1) and (CStr(prim_suporte) <> "0") and (CStr(prim_suporte) <> "") then
			url = url & "&suporte=" & prim_suporte
		end if
		
		'Por algum motivo no TCU, o comando "Response.Redirect" não está funcionando neste ponto
		if (global_numero_serie = 3621) then
			Response.Write "<script type='text/javascript'>"
			Response.Write " document.location='" & url & "';"
			Response.Write "</script>"
		else				
			Response.Redirect url
		end if
	end if
	
%>
</div>
</body>
</html>