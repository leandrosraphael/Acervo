<!DOCTYPE html>
<% 
	htmlClass = ""
	if Request.Cookies("contraste") = "1" then
		htmlClass = "class='contraste'"
	end if
%>
<html <%=htmlClass%>>
<head>
<% sDiretorioArq="asp" %>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<script type='text/javascript'>
<!--
function Trim(TRIM_VALUE){
	if(TRIM_VALUE.length < 1) {
		return "";
	}
	TRIM_VALUE = RTrim(TRIM_VALUE);
	TRIM_VALUE = LTrim(TRIM_VALUE);
	if(TRIM_VALUE == "") {
		return "";
	} else {
		return TRIM_VALUE;
	}
} //End Function

function RTrim(VALUE){
	var w_space = String.fromCharCode(32);
	var v_length = VALUE.length;
	var strTemp = "";
	if(v_length < 0) {
		return	"";
	}
	var iTemp = v_length -1;

	while(iTemp > -1) {
		if(VALUE.charAt(iTemp) == w_space) {
		} else {
			strTemp = VALUE.substring(0,iTemp +1);
			break;
		}
		iTemp = iTemp-1;
	} //End While
	return strTemp;
} //End Function

function LTrim(VALUE){
	var w_space = String.fromCharCode(32);
	if(v_length < 1) {
		return "";
	}
	var v_length = VALUE.length;
	var strTemp = "";

	var iTemp = 0;

	while(iTemp < v_length) {
		if(VALUE.charAt(iTemp) == w_space) {
		} else {
			strTemp = VALUE.substring(iTemp,v_length);
			break;
		}
		iTemp = iTemp + 1;
	} //End While
	return strTemp;
} //End Function


function ValidaCamposFormulario() {
	var permite_envio_email = true;

	if (global_numero_serie == 3621) {
		permite_envio_email = ValidarCamposObrigatorios();
	}

	if ((permite_envio_email) && (document.frm_sels.sels_email_usuario.value != "")) {
		permite_envio_email = validateEmail(frm_sels.sels_email_usuario.value,
			'<%=getTermo(global_idioma, 9874, "O e-mail do remetente é inválido.", 0)%>');
	}

	if (permite_envio_email) {
		permite_envio_email = validateEmail(document.frm_sels.sels_email.value,
			'<%=getTermo(global_idioma, 9875, "O e-mail do destinatário é inválido.", 0)%>');
	}

	if (permite_envio_email) {
		document.frm_sels.submit();
	}
	else {
		return false;
		document.frm_sels.focus();
	}
}

function validateEmail(email, mensagem_alerta)
{
	email = email.replace(",",";");
	var array_email=email.split(";");
	var email_num=0;
	var invalido = 0;
	while (email_num < array_email.length)
	{
		array_email[email_num] = Trim(array_email[email_num]);
		//alert(array_email[email_num]);
		if (typeof(array_email[email_num]) != "string") {
			invalido+=1;
		}else if (!array_email[email_num].match(/^[A-Za-z0-9]+([_.-][A-Za-z0-9]+)*@[A-Za-z0-9]+([_.-][A-Za-z0-9]+)*\.[A-Za-z0-9]{2,4}$/)) {
			invalido+=1;
		}
		email_num+=1;
	}
	

	if (invalido > 0) {
		alert(mensagem_alerta);
		email.focus();
		return false;
	}else{
		return true;
	}
}

function ValidarCamposObrigatorios()
{
	var stermo = "";
	var bExibeErro = false;

	if (document.frm_sels.sels_nome_usuario.value == "") {
		stermo = '<%=getTermo(global_idioma, 6660, "O campo %s é obrigatório.", 0)%>';
		stermo = stermo.replace(/%s/g, '<%=getTermo(global_idioma, 94, "Nome", 0)%>');
		bExibeErro = true;
		campo = document.frm_sels.sels_nome_usuario;
	}
	else if (document.frm_sels.sels_email_usuario.value == "") {
		stermo = '<%=getTermo(global_idioma,  9874, "O e-mail do remetente é inválido.", 0)%>';
		bExibeErro = true;
		campo = document.frm_sels.sels_email_usuario;
	}

	if (bExibeErro) {
		alert(stermo);
		campo.focus();
		return false;
	}
	else {
		return true;
	}

}

//-->
</script>
<title>::<%=getTermo(global_idioma, 963, "Minha seleção", 0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />
<%'--------------------------------------------- 
if config_css_estatico = 1 then 
%>
	<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
<%
else 
%>
	<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
<% 
end if 
%>
<link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
</head>

<%'--------------------------------------------- 
if Request.QueryString("acao") = "conferir" then 
%>
	<body class="popup" onload="document.frm_sels.sels_nome_usuario.focus();">
<% 
else 
%>
	<body class="popup">
<%
end if
%>
<script type='text/javascript'>parent.fechaLoadingPopup();</script>
<div class="centro">
<%
iIndexSrv = IntQueryString("Servidor", 1)
iCodigoFavorito = Request.Form("codigoFavorito") 
'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

if Request.QueryString("acao") = "enviar" then
	
	EmailSel = replace(request.Form("sels_email"),",",";")
	TipoSel = request.Form("sels_tipo")
	NomeUsuario = request.Form("sels_nome_usuario")
	EmailUsuarioCorpo = request.Form("sels_email_usuario")
	EmailAssunto = request.Form("sels_assunto")
	EmailMensagem = request.Form("sels_mensagem")
	
	if (iCodigoFavorito > 0) then
		if TipoSel = "lista" then
			iTipo = 7	
		else
			iTipo = 1
		end if
		bOk = true
	else
		Sessao = "mysel"&iIndexSrv
		if Session(Sessao) = "" then
			Response.Write "<span style='color: red'>"&getTermo(global_idioma, 1474, "Sessão expirada", 0)&"</span>"
			bOk = false
		else
			mysel = Session(Sessao)		
			arMySel = split(mysel,",")
			sMsgInfo = ""						
			bOk = true
			if TipoSel = "lista" then
				iTipo = 2	
			else
				iTipo = 1
		
				iNumLeg = 0
				iNumOutros = 0
		
				for i = lbound(arMySel) to ubound(arMySel)
				
					arTemp   = split(arMySel(i),".")
					TipoObra = arTemp(1)
				
					if(cInt(TipoObra) <> 2) then						
						iNumOutros = iNumOutros + 1
					else
						iNumLeg = iNumLeg + 1
					end if
				Next
				
				if(iNumLeg = 0)and(iNumOutros = 0)then
					Response.Write "<br />"&getTermo(global_idioma, 1475, "Nenhuma seleção no momento", 0)

				elseif(iNumLeg > 1)and(iNumOutros > 0)then
					sMsgInfo = "<br /><br />"&getTermo(global_idioma, 1476, "Obs: As legislações selecionadas não foram incluídas no e-mail, pois não está disponível a referência bibliográfica para este tipo de material.", 0)&"<br />"&getTermo(global_idioma, 1477, "Utilize a opção ""lista de materiais"" para enviar todos os registros selecionados.", 0)
				elseif(iNumLeg > 0)and(iNumOutros = 0)then		
					Response.Write "<script type='text/javascript'>alert('"&getTermo(global_idioma, 1478, "Não está disponível a referência bibliográfica para este tipo de material.", 0)&"\n"&getTermo(global_idioma, 1479, "Utilize a opção ""lista de materiais"".", 0)&"');</script>"
					Response.Write "<script type='text/javascript'>parent.fechaPopup();</script>" 
				elseif(iNumLeg = 1)and(iNumOutros > 0)then		
					sMsgInfo = "<br /><br />"&getTermo(global_idioma, 1480, "Obs: A legislação selecionada não foi incluída no e-mail, pois não está disponível a referência bibliográfica para este tipo de material.", 0)&"<br />"&getTermo(global_idioma, 1477, "Utilize a opção ""lista de materiais"" para enviar todos os registros selecionados.", 0)
				end if	
			end if
		end if			 
	end if

	if (bOk) then	
		if(Session("bib_usuario") <> "")then
			iBib = Session("bib_usuario")
		else
			iBib = 1
		end if

		if(Session("codigo_usuario") <> "")then
			iUsuario = Session("codigo_usuario")
		else
			iUsuario = 0
		end if
		
		'Chamada a rotina que envia os e-mails		
		On Error Resume Next
		SET ROService = ROServer.CreateService("Web_Consulta")
		Set Msg = ROService.EnviaSelecaoPorEmail(EmailSel, EmailAssunto, EmailMensagem, mysel, iTipo, iBib, iCodigoFavorito, global_idioma, iUsuario, NomeUsuario, EmailUsuarioCorpo)
		SET ROService = nothing		
			
		if (Msg.Result) then
			sMsgOk = getTermo(global_idioma, 1482, "E-mail enviado com sucesso.", 0)& sMsgInfo & Msg.sMsg
		else						
			sMsgOk = getTermo(global_idioma, 1481, "Erro ao enviar o e-mail.", 0) & Msg.sMsg
		end if
		%>
		<body class="popup">
		<script type="text/javascript">parent.fechaLoadingPopup();</script>
		<br />
		<div class="centro">
			<b><%=sMsgOk%></b>
			<br /><br /><br />
			<a href="javascript:parent.fechaPopup('asp');" class='link_serv'><%=getTermo(global_idioma, 4100, "fechar janela", 2)%></a>
		</div>
		</body>
		<%
	end if
	
elseif Request.QueryString("acao") = "conferir" then

	'Pega a variável de sessão com o e-mail do usuário
	sEmailUsu = Session("Email_Usuario")
	
	Response.Write "<div style='margin-top: 15px; margin-bottom: 10px;'>"&getTermo(global_idioma, 1484, "Enviar e-mail com", 0)&":</div><form name='frm_sels' action='envia_sels.asp?acao=enviar&Servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iIdioma="&global_idioma&"' method='post'>"
	Response.Write "<table style='margin: auto'>"
	Response.Write "	<tr>"
	Response.Write "		<td colspan='2' class='esquerda'>"
	Response.Write "		    <div style='padding-bottom: 4px'>"
	Response.Write "			<input type=radio name=sels_tipo value='lista' checked>&nbsp;" & getTermo(global_idioma, 1485, "Lista de materiais", 0)
	Response.Write "			</div>"
	Response.Write "			<input type=radio name=sels_tipo value='refbib'>&nbsp;" & getTermo(global_idioma, 792, "Referência bibliográfica", 0)
	Response.Write "		</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	
	Response.Write "<br />"
	
	if (global_numero_serie = 3621) then 'TCU
		sCampoObrigatorio = "*"
		sEmailDe = sEmailUsu
		sEmailUsu = ""
	else
		sEmailDe = ""
		sCampoObrigatorio = ""
	end if

	Response.Write "<table style='margin: auto'>"
	Response.Write "	<tr style='height: 26px'><td class='esquerda' style='width: 70px;'>" & getTermo(global_idioma, 94, "Nome", 0) & sCampoObrigatorio & "</td> <td>"
	Response.Write "	<input type='text' class='input354' name='sels_nome_usuario' maxsize='90'/></td></tr>"
	Response.Write "	<tr style='height: 26px'><td class='esquerda' style='width: 70px;'>" & getTermo(global_idioma, 6018, "De", 0) & sCampoObrigatorio & "</td> <td>"
	Response.Write "	<input type='text' class='input354' name='sels_email_usuario' value='" & sEmailDe & "' maxsize='90'/></td></tr>"
	Response.Write "	<tr style='height: 26px'><td class='esquerda' style='width: 70px;'>" & getTermo(global_idioma, 1486, "Para", 0) &  "*</td> <td>"
	Response.Write "	<input type='text' class='input354' name='sels_email' value='" & sEmailUsu & "' maxsize='90'/></td></tr>"
	Response.Write "	<tr class='separador_campo' style='height: 26px'><td style='width: 70px;'></td><td class='esquerda font_11px'>*" & getTermo(global_idioma, 9876, "Campo(s) obrigatório(s).", 0) & "</td></tr>"

	Response.Write "	<tr style='height: 26px'><td class='esquerda'>" & getTermo(global_idioma, 72, "Assunto", 0) &  "</td> <td>"
	if (Request.QueryString("codigoFavorito") <> "0") then
		Response.Write "	<input type='text' class='input354' name=sels_assunto value='" & getTermo(global_idioma, 8321, "Lista de favoritos", 0) & "' maxsize='90'></td></tr>"
	else
		Response.Write "	<input type='text' class='input354' name=sels_assunto value='" & getTermo(global_idioma, 963, "Minha seleção", 0) & "' maxsize='90'></td></tr>"
	end if
	Response.Write "	<tr style='height: 58px'><td class='esquerda'>" & getTermo(global_idioma, 370,"Mensagem", 0) &  "</td> <td>"
	Response.Write "	<textarea name='sels_mensagem' style='width: 356px; height: 70px;'></textarea></td></tr>"
	Response.Write "</table>"
	Response.Write "<input type=hidden name=codigoFavorito value="&Request.QueryString("codigoFavorito") &">"
	Response.Write "<br /><input type=hidden name=sub_sels value='sim'><input type=button onClick='ValidaCamposFormulario();' value='" & getTermo(global_idioma, 4, "Confirmar", 0) & "'>&nbsp;&nbsp;"
	Response.Write "<input type='button' onclick='parent.fechaPopup();' value='" & getTermo(global_idioma, 5, "Cancelar", 0) & "'/></form>"
end if
%>
</div>
</body>
</html>