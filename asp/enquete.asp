<% Response.Buffer=True %>
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
<title>::Enquete</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />
<% if config_css_estatico = 1 then %>
	<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
<% else %>
	<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
<% end if %>
<script src="../scripts/jquery/jquery-3.2.1.min.js"></script>

<script type="text/javascript" src="../scripts/jquery/jquery-ui.min.js"></script>
<script type="text/javascript" src="../scripts/jquery.multiselect.js"></script>
<script type="text/javascript" src="../scripts/combo.init.js"></script>

<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
<script type="text/javascript">
	function limitarCampo(campo, limite) {
		if (limite > 0) {
			if (campo.value.length > limite) {
				campo.value = campo.value.substring(0, limite);
			}
		}
	}

	function validaFormulario() {
		var questoes = document.getElementsByTagName('div');

		for (i = 0; i < questoes.length; i++) {
			if (questoes[i].className == "div_questao") {
				var ques_fechada = questoes[i].getAttribute('fechada');
				var num_questao = questoes[i].getAttribute('questao');
				var tipo_escolha = questoes[i].getAttribute('tipo_escolha');
				var tipo_apres = questoes[i].getAttribute('tipo_apresentacao');

				// Valida se a questao é fechada ou não
				if ((ques_fechada != undefined) && (ques_fechada != null) &&
					(num_questao != undefined) && (num_questao != null) &&
					(tipo_escolha != undefined) && (tipo_escolha != null) &&
					(tipo_apres != undefined) && (tipo_apres != null)) {
					//Questão fechada
					if (ques_fechada == '1') {
						//Seleção única (Radio ou Combo)
						if (tipo_escolha == 0) {
							//Radio
							if (tipo_apres == 0) {
								var comp = null;
								var resposta = document.getElementsByName('questao' + num_questao + '_resposta');
								if ((resposta != undefined) && (resposta != null)) {
									var check = false;

									for (radio = 0; radio < resposta.length; radio++) {
										if (radio == 0) {
											comp = resposta[radio];
										}
										if ((resposta[radio].checked) && (resposta[radio].value > 0)) {
											check = true;
										}
									}

									if (!check) {
										var msg = getTermo(global_frame.iIdioma, 6034, 'Responda corretamente a questão %s.', 0);
										msg = msg.replace('%s', num_questao);
										alert(msg);
										if (comp != null)
											comp.focus();
										return false;
									}
								}
							}
							//Combo
							else {
								var resposta = document.getElementById('questao' + num_questao + '_resposta');
								if ((resposta != undefined) && (resposta != null)) {
									if (resposta.value <= 0) {
										var msg = getTermo(global_frame.iIdioma, 6034, 'Responda corretamente a questão %s.', 0);
										msg = msg.replace('%s', num_questao);
										alert(msg);
										resposta.focus();
										return false;
									}
								}
							}
						}
						//Seleção Múltipla (Check-box)
						else if (tipo_escolha == 1) {
							//Número de alternativas da questão
							var qtde_resp = questoes[i].getAttribute('num_respostas');
							if ((qtde_resp != undefined) && (qtde_resp != null)) {
								var comp = null;
								var qtde_sel = 0;

								//Conta quantas alterativas foram selecionadas
								for (resp = 1; resp <= qtde_resp; resp++) {
									var resposta = document.getElementById('questao' + num_questao + '_resposta' + resp);
									if (resp == 1) {
										comp = resposta;
									}
									if ((resposta != undefined) && (resposta != null)) {
										if ((resposta.checked) && (resposta.value > 0)) {
											qtde_sel++;
										}
									}
								}

								if (qtde_sel <= 0) {
									var msg = getTermo(global_frame.iIdioma, 6034, 'Responda corretamente a questão %s.', 0);
									msg = msg.replace('%s', num_questao);
									alert(msg);
									if (comp != null)
										comp.focus();
									return false;
								}
							}
						}
					}
					//Questão aberta
					else {
						//Valida o preenchimento de questões do tipo Memo
						var resposta = document.getElementById('questao' + num_questao + '_resposta');
						if ((resposta != undefined) && (resposta != null)) {
							if (resposta.value.length <= 0) {
								var msg = getTermo(global_frame.iIdioma, 6034, 'Responda corretamente a questão %s.', 0);
								msg = msg.replace('%s', num_questao);
								alert(msg);
								resposta.focus();
								return false;
							}
						}
					}
				}
			}
		}

		document.frm_enquete.submit();
	}
</script>

<link href="../scripts/css/jquery-ui-1.8.18.datepicker.css" rel="stylesheet" />

<link href="../scripts/css/jquery.multiselect.css" rel="stylesheet" />
<link href="../scripts/css/multiselect-custom.css" rel="stylesheet" />
<link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />

</head>
<body class="popup" onload="parent.fechaLoadingPopup();">
<p class="centro">
<table class="centro tab_enquete">
<tr>
<td class="td_center_top">
<%
if config_multi_servbib = 1 then
	iIndexSrv = Session("Servidor")

	if iIndexSrv = "" then
		iIndexSrv = 1
	end if

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
end if

iEnquete = IntQueryString("enquete", 0)
iCodUsu  = Session("codigo_usuario")

if (Request.QueryString("acao") = "responder") then
	iPasso  = CLng(Request.Form("passo"))
	iGrupos = IntForm("grupos", 0)

	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	
	sXMLQuestao = ROService.QuestoesEnquete(iEnquete)

	SET ROService = nothing
	TrataErros(1)
	
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	
	sFormulario = ""
	iIDQuestao = 0
	
	xmlDoc.async = False
	xmlDoc.loadxml sXMLQuestao
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "grupos" then
		identificacao = xmlRoot.attributes.getNamedItem("identificacao").value
		
		For Each xmlGrupo In xmlRoot.childNodes	
			if (sFormulario <> "") then
				sFormulario = sFormulario & "<br /><br />"
			end if
			
			desc_grupo = xmlGrupo.attributes.getNamedItem("descricao").value
			sFormulario = sFormulario & "<p class='centro'><b>" & desc_grupo & " </b></p>"
	
			For Each xmlQuestoes In xmlGrupo.childNodes
				'*************************************
				'QUESTÃO
				'*************************************
				if xmlQuestoes.nodeName  = "questao" then
					iIDQuestao = iIDQuestao + 1
					sFormulario = sFormulario & "<br /><br />"
					
					cod_questao = xmlQuestoes.attributes.getNamedItem("codigo").value
					ques_fechada = xmlQuestoes.attributes.getNamedItem("fechada").value
					char_resp = xmlQuestoes.attributes.getNamedItem("caracteres").value
					desc_questao = xmlQuestoes.attributes.getNamedItem("descricao").value
					
					sFormulario = sFormulario & CStr(iIDQuestao) & ") "	& desc_questao & "<br /><br />"
					sFormulario = sFormulario & "<input type='hidden' name='questao" & iIDQuestao & "_fechada' value='" & ques_fechada & "'/>"
					sFormulario = sFormulario & "<input type='hidden' name='questao" & iIDQuestao & "_codigo' value='" & cod_questao & "'/>"
					
					'Questão fechada
					if (ques_fechada = "1") then
						For Each xmlResposta In xmlQuestoes.childNodes
							if (xmlResposta.nodeName = "respostas") then
								tipo_escolha = xmlResposta.attributes.getNamedItem("tipo").value
								tipo_apres = xmlResposta.attributes.getNamedItem("apresentacao").value
								num_resp = xmlResposta.attributes.getNamedItem("num_alternativas").value
								
								sFormulario = sFormulario & "<input type='hidden' name='questao" & iIDQuestao & "_tipo_escolha' value='" & tipo_escolha & "'/>"
								sFormulario = sFormulario & "<input type='hidden' name='questao" & iIDQuestao & "_num_resp' value='" & num_resp & "'/>"
								sFormulario = sFormulario & "<div class='div_questao' questao='" & iIDQuestao & "' fechada='1' tipo_escolha='" & tipo_escolha & "' tipo_apresentacao='" & tipo_apres & "' num_respostas='" & num_resp & "' name='div_questao" & iIDQuestao & "' id='div_questao" & iIDQuestao & "'>"
								
								'Seleção única
								if (tipo_escolha = "0") then
									'Combo
									if (tipo_apres = "1") then
										sFormulario = sFormulario & "<select class='select-enquete styled_combo' name='questao" & iIDQuestao & "_resposta' id='questao" & iIDQuestao & "_resposta'>"
									end if
								end if
								
								seq_resp = 0
							
								For Each xmlAlternativa In xmlResposta.childNodes
									'*************************************
									'ALTERNATIVA
									'*************************************
									if (xmlAlternativa.nodeName = "alternativa") then
										seq_resp = seq_resp + 1									
									
										cod_alternativa = xmlAlternativa.attributes.getNamedItem("codigo").value
										desc_alternativa = xmlAlternativa.text
										
										'Seleção única
										if (tipo_escolha = "0") then
											'Radio
											if (tipo_apres = "0") then
												sFormulario = sFormulario & "<input type='radio' name='questao" & iIDQuestao & "_resposta' id='questao" & iIDQuestao & "_resposta" & seq_resp & "' value='" & cod_alternativa & "'/>"
												sFormulario = sFormulario & "<label for='questao" & iIDQuestao & "_resposta" & seq_resp & "'>" & desc_alternativa & "</label><br />"
											'Combo
											else
												sFormulario = sFormulario & "<option value='" & cod_alternativa & "'>" & desc_alternativa & "</option>"
											end if
										'Seleção múltipla
										else
											sFormulario = sFormulario & "<input type='checkbox' name='questao" & iIDQuestao & "_resposta' id='questao" & iIDQuestao & "_resposta" & seq_resp & "' value='" & cod_alternativa & "'>"
											sFormulario = sFormulario & "<label for='questao" & iIDQuestao & "_resposta" & seq_resp & "'>" & desc_alternativa & "</label><br />"
										end if																
									end if
								Next
								
								'Seleção única
								if (tipo_escolha = "0") then
									'Combo
									if (tipo_apres = "1") then
										sFormulario = sFormulario & "</select><br />"
									end if
								end if
								
								sFormulario = sFormulario & "</div>"
							end if
						Next
					'Questão aberta
					else
						sFormulario = sFormulario & "<div class='div_questao' questao='" & iIDQuestao & "' fechada='0' tipo_escolha='0' tipo_apresentacao='0' num_respostas='0' name='div_questao" & iIDQuestao & "' id='div_questao" & iIDQuestao & "'>"
						sFormulario = sFormulario & "<textarea name='questao" & iIDQuestao & "_resposta' id='questao" & iIDQuestao & "_resposta' style='width: 600px; height: 60px;' onKeyDown='limitarCampo(this,"&char_resp&");' onKeyUp='limitarCampo(this,"&char_resp&");'></textarea>"
						sFormulario = sFormulario & "</div><br />"
					end if			
				end if
			Next
		Next
		
		'Sem Identificação
		if (CInt(identificacao) = 0) then
			sFormulario = sFormulario & "<input type='hidden' name='questao_identificacao_anonima' id='questao_identificacao_anonima' value='1'>"
		'Identificação opcional
		elseif (CInt(identificacao) = 2) then
			sFormulario = sFormulario & "<br /><br />"
			
			sFormulario = sFormulario & "<input type='checkbox' name='questao_identificacao_anonima' id='questao_identificacao_anonima' value='1'/>"
			sFormulario = sFormulario & "<label for='questao_identificacao_anonima'>" & getTermo(global_idioma, 6036, "Não permitir a minha identificação", 0) & "</label><br />"
		end if
	end if
	
	Set xmlDoc = nothing
	Set xmlRoot = nothing
%>
	<form action="enquete.asp?enquete=<%=iEnquete%>&acao=gravar&Servidor=<%=iIndexSrv%>" method="post" name="frm_enquete" id="frm_enquete"> 	
	    <p class="centro">
		    <div style="height: 400px; max-height: 400px; overflow: auto; text-align:justify;">
			    <%=sFormulario%>
                <br />
		    </div>
            <input class="button_termo" type="button" onClick="validaFormulario();" value="Confirmar" id="button3" name="button1"/>
		    <input class="button_termo" type="button" onClick="parent.fechaPopup();" value="<%=getTermo(global_idioma, 5, "Cancelar", 0)%>" id="button4" name="button2"/>
            <input type="hidden" name="grupos" value="<%=iGrupos%>"/>
            <input type="hidden" name="questoes" value="<%=iIDQuestao%>"/>
	    </p>
    </form>
	
	</td>
	</tr>
	<tr style="height:30px">
		<td class="centro">
			
		</td>
	</tr>
<%
elseif (Request.QueryString("acao") = "gravar") then
	iQuestoes = Request.Form("questoes")
	iRepostaAnonima = CStr(Request.Form("questao_identificacao_anonima"))

	sXML = "<?xml version=""1.0"" encoding=""windows-1252""?>"
	
	if (iRepostaAnonima = "1") then
		sXML = sXML & "<questoes resposta_anonima=""1"">"
	else
		sXML = sXML & "<questoes resposta_anonima=""0"">"
	end if
	
	for iQuestao = 1 to iQuestoes	
	
    	iFechada = Request.Form("questao" & iQuestao & "_fechada")

		iCodigo  = Request.Form("questao" & iQuestao & "_codigo")
	
		sXML = sXML & "<questao fechada=""" & iFechada & """ codigo=""" & iCodigo & """>"
		
		'Resposta fechada
		if (iFechada = 1) then
			iTipoEscolha = Request.Form("questao" & iQuestao & "_tipo_escolha")

			'Seleção única
			if (iTipoEscolha = 0) then
				sXML = sXML & Request.Form("questao" & iQuestao & "_resposta")
			'Seleção Múltipla
			else
				sResp = ""
				For Each resposta In Request.Form("questao" & iQuestao & "_resposta")
					if (sResp <> "") then
						sResp = sResp & ","
					end if
					sResp = sResp & resposta
				Next
				
				sXML = sXML & sResp
			end if
		'Resposta aberta
		else
			sXML = sXML & Request.Form("questao" & iQuestao & "_resposta")
		end if
		sXML = sXML & "</questao>"
	Next
	
	sXML = sXML & "</questoes>"

	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	
	sXMLRespostas = ROService.GravaEnquete(iEnquete, iCodUsu, sXML)

	SET ROService = nothing
	TrataErros(1)
	
	Response.Write "<script type='text/javascript'>"
	Response.Write "  if (parent.parent.hiddenFrame != null) {"
	Response.Write "    Hiddenfrm = parent.parent.hiddenFrame;"
	Response.Write "	Hiddenfrm.popup_refresh = true;"
	Response.Write "  }"
	Response.Write "</script>"
	
	Set xmlDoc = CreateObject("Microsoft.xmldom")	
	xmlDoc.async = False
	xmlDoc.loadxml sXMLRespostas
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "resposta" then
		For Each xmlResposta In xmlRoot.childNodes	
			if xmlResposta.nodeName  = "titulo" then
				enq_agradecimento_titulo = xmlResposta.text
			end if
			if xmlResposta.nodeName  = "mensagem" then
				enq_agradecimento = xmlResposta.text
			end if
		Next
	end if
	
	Set xmlRoot = nothing
	Set xmlDoc = nothing

%>   
	<P class="centro">
		<b><%=enq_agradecimento_titulo%></b>
		<br /><br />
		<div style="height: 100%; max-height: 420px; overflow: auto; text-align:justify">
		<%=replace(enq_agradecimento,chr(10),"<br />")%>
		</div>    
	</p>
	
	</td>
	</tr>
	<tr style="height:30px">
		<td class="centro">
			<input class="button_termo" type="button" onClick="parent.fechaPopup();" value="<%=getTermo(global_idioma, 220, "Fechar", 0)%>" id="button2" name="button2"/>
		</td>
	</tr>
<%
else
	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	
	sXMLEnquete = ROService.InfoEnquete(iEnquete, iCodUsu, global_idioma)
	
	SET ROService = nothing
	TrataErros(1)
	
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	
	xmlDoc.async = False
	xmlDoc.loadxml sXMLEnquete
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "enquete" then
		For Each xmlCampo In xmlRoot.childNodes
			'*************************************
			'RESPONDIDO
			'*************************************
			if xmlCampo.nodeName  = "respondido" then
				enq_respondido = xmlCampo.text
			end if
			'*************************************
			'TITULO
			'*************************************
			if xmlCampo.nodeName  = "titulo" then
				enq_titulo = xmlCampo.text
			end if
			'*************************************
			'INSTRUÇÃO
			'*************************************
			if xmlCampo.nodeName  = "instrucao" then
				enq_mensagem = xmlCampo.text
			end if	
			'*************************************
			'GRUPOS
			'*************************************
			if xmlCampo.nodeName  = "grupos" then
				enq_grupos = xmlCampo.text
			end if			
		Next
	end if	
	
	Set xmlDoc = nothing
	Set xmlRoot = nothing
	
	%>
	
	<p class="centro">
		<b><%=enq_titulo%></b>
		<br /><br />
		<div style="height: 100%; max-height: 420px; overflow: auto; text-align:justify">
		<%=replace(enq_mensagem,chr(10),"<br />")%>
		</div>    
	</p>
	
	</td>
	</tr>
	<tr style="height:30px">
		<td class="centro">
        	<form action="enquete.asp?enquete=<%=iEnquete%>&acao=responder&Servidor=<%=iIndexSrv%>" method="post" name="frm_enquete" id="frm_enquete">
			<% if (enq_respondido <> "1") then %>
				<input class="button_termo" type="button" onClick="submit();" value="Responder" id="button1" name="button1">
			<% end if %>
			<input class="button_termo" type="button" onClick="parent.fechaPopup();" value="<%=getTermo(global_idioma, 5, "Cancelar", 0)%>" id="button2" name="button2"/>
            <input type="hidden" name="grupos" value="<%=enq_grupos%>"/>
            </form>
		</td>
	</tr>
<% end if %>
</table>
</p>
</body>
</html>