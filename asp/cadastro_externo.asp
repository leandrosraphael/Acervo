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
<%
global_idioma = Request.QueryString("idioma")
%>
<title>::<%=getTermo(global_idioma, 9298, "Cadastrar", 0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />

<% if config_css_estatico = 1 then %>
	<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
<% else %>
	<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
<% end if 
 
	iIndexSrv = 1 

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	bExibirTermoAceite = false
	On Error Resume Next
					   
	SET ROService = ROServer.CreateService("Web_Consulta")
	bExibirTermoAceite = ROService.ExibirTermoAceite
	Set ROService = nothing
	TrataErros(1)

	sXMLEmpresas = ""
	if (global_numero_serie = 3156) then
		On Error Resume Next
					   
		SET ROService = ROServer.CreateService("Web_Consulta")
		sXMLEmpresas = ROService.DadosTabela(roTAB_EMPRESA,"",0, repositorio_institucional)
		Set ROService = nothing
		TrataErros(1)
	end if

	%>

<link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
<link href="../inc/cadastro_externo.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
<script src="../scripts/jquery/jquery-3.2.1.min.js"></script>

<script type="text/javascript">
	function FocoLogin() {
		try {
			document.login.codigo.focus();
			$("div#mensagem").hide();
		} catch (er) { }
	}

	function validarFormCadastroExterno() {
		var nome = $("#nome").val();
		var email = $("#email").val();
		var senha = $("#senha").val();
		var confirmacao_senha = $("#confirmacao_senha").val();
		var empresa = $("#codigo_empresa").val();
		var nome_empresa = $("#nome_empresa").val();
		var aceite = document.getElementById("termo_aceite");

		if (nome.length == 0) {
			alert('<%= Format(getTermo(global_idioma, 320, "O campo ""%s"" deve ser preenchido.", 0), getTermo(global_idioma, 94, "Nome", 2))%>');
			return false;
		}

		if (email.length == 0) {
			alert('<%= Format(getTermo(global_idioma, 320, "O campo ""%s"" deve ser preenchido.", 0), getTermo(global_idioma, 129, "E-mail", 2))%>');
			return false;
		}

		if (senha.length == 0) {
			alert('<%= Format(getTermo(global_idioma, 320, "O campo ""%s"" deve ser preenchido.", 0), getTermo(global_idioma, 2, "Senha", 2))%>');
			return false;
		}

		if (confirmacao_senha.length == 0) {
			alert('<%= Format(getTermo(global_idioma, 320, "O campo ""%s"" deve ser preenchido.", 0), getTermo(global_idioma, 9295, "Confirmação de senha", 2))%>');
			return false;
		}

		if (confirmacao_senha != senha) {
			alert('<%=getTermo(global_idioma, 4237, "A senha não confere.", 0) & " " & getTermo(global_idioma, 4236, "Digite novamente.", 0)%>');
			return false;
		}

	<% if (bExibirTermoAceite) then %>
		if (!aceite.checked) {
			alert('<%=getTermo(global_idioma, 9286, "É necessário aceitar os termos e condições.", 0)%>');
			return false;
		}
	<%end if %>

	<%
	if (global_numero_serie = 3156) then
				%>
		if (!empresa || empresa.length == 0) {
			alert('<%=getTermo(global_idioma, 47, "A empresa do usuário deve ser preenchida.", 0)%>');
			return false;
		}

		<% if (global_cfg_cadastrar_empresa_usuario_externo = "1") then %>
			if (empresa == "-1" && nome_empresa.length == 0) {
				 alert('<%=getTermo(global_idioma, 47, "A empresa do usuário deve ser preenchida.", 0)%>');
				return false;
			}
	   <% end if %>
	<%
	end if
	%>

        cadastrar_usuario_externo($("form[name='formUsuarioExterno']").serialize(),<%=global_idioma%>);
		return true;
	}
</script>
</head>
<body class="popup centro" onload="FocoLogin(); parent.fechaLoadingPopup();">
	<div id="conteudo-principal">
		<form class="formLogin" name="formUsuarioExterno">
			<div class="aviso">
				<p>	
					<%=getTermoHtml(global_idioma, 9303, "Por favor, informe os seus dados", 0)%>
				</p>
			</div>
			<div class="divForm">
				<label for="nome"><%=getTermo(global_idioma, 94, "Nome", 0)%></label>
				<input type="text" class="input" name="nome" id="nome" maxlength="252" onKeyPress="return validaTeclaLogin(this,event,2)"/>
			</div>
			<div class="divForm">
				<label for="email"><%=getTermo(global_idioma, 129, "E-mail", 0)%></label>
				<input type="text" class="input" name="email" id="email" maxlength="252" onKeyPress="return validaTeclaLogin(this,event,2)"/>
			</div>
			<div class="divForm">
				<label for="senha"><%=getTermo(global_idioma, 2, "Senha", 0)%></label>
				<input type="password" name="senha" id="senha" class="input" maxlength="30" onKeyPress="return validaTeclaLogin(this, event,0)" />
			</div>
			<div class="divForm">
				<label for="confirmacao_senha"><%=getTermo(global_idioma, 9295, "Confirmação de senha", 0)%></label>
				<input type="password" name="confirmacao_senha" id="confirmacao_senha" class="input" maxlength="30" onKeyPress="return validaTeclaLogin(this, event,0)" />
			</div>
		
			<% if (global_numero_serie = 3156) then 
					if (left(sXMLEmpresas,5) = "<?xml") then %>
						<div class="divForm">
							<label for="codigo_empresa"><%=getTermo(global_idioma, 120, "Empresa", 0)%></label>
							<%if (global_cfg_cadastrar_empresa_usuario_externo = "1") then %>
								<select name="codigo_empresa" id="codigo_empresa" onchange="permitirOutroCadastro(this,event);"/>
							<% else %>
								<select name="codigo_empresa" id="codigo_empresa" />
							<%end if
						
								'Processa o XML
								Set xmlDoc = CreateObject("Microsoft.xmldom")
								xmlDoc.async = False
								xmlDoc.loadxml sXMLEmpresas
								Set xmlRoot = xmlDoc.documentElement

								'Verifica se a tabela foi encontrada e se possui registros
								if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
									For Each xmlTabela In xmlRoot.childNodes
										if xmlTabela.nodeName  = "REGISTRO" then
											sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
											sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
											Response.Write "<option value="&sCod&">"&sDesc&"</option>"
										end if
									Next
								end if
								if (global_cfg_cadastrar_empresa_usuario_externo = "1") then
								   Response.Write "<option value='-1'>" & getTermo(global_idioma, 1867, "Outra", 0) & " - " & getTermo(global_idioma, 221, "Incluir", 0) & "</option>"
								end if
								%>
							</select>
						</div>

						<% if (global_cfg_cadastrar_empresa_usuario_externo = "1") then %>
							<% if (xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0") then %>
								<div class="divForm" id="area_cadastro_empresa" style="visibility: hidden">
							<% else %>
								<div class="divForm" id="area_cadastro_empresa">
							<% end if %>
								<label for="nome"><%=getTermo(global_idioma, 1867, "Outra", 0)%></label>
								<input type="text" id="nome_empresa" name="nome_empresa" maxlength="30"/>
							</div>
						<% else %>
							<input type="hidden" id="nome_empresa" name="nome_empresa" value="" />
						<%end if%>
					<%end if
			   else %>
					<input type="hidden" id="codigo_empresa" name="codigo_empresa" value="0" />
					<input type="hidden" id="nome_empresa" name="nome_empresa" value="" />
			<% end if %>

			<% if (bExibirTermoAceite) then %>
				<div class="check-termo-aceite">
					<input type="checkbox" id="termo_aceite" name="termo_aceite" value="1" /><span><%=getTermo(global_idioma, 9296, "Concordo com os termos e condições.", 0)%></span>
					<p><a href="termo_aceite_usuario_externo.asp?idioma=<%=global_idioma%>" target="_blank"><%=gettermo(global_idioma, 9297, "Clique aqui para visualizar o termo.", 0) %></a></p>
				</div>
			<% else %>
				<input type="hidden" name="termo_aceite" value="1"/>
			<% end if %>

			<% if config_multi_servbib = 1 then %>
				<div class="divForm">
					<label><%=getTermo(global_idioma, 3, "Biblioteca", 0)%></label>
					<select name="servidor">
					<%
						iBibSel = IntQueryString("Servidor", 1)
				
						iNumAba = Servidores.ServList.Count
						'Cria uma aba para cada servidor
						For i = 1 to iNumAba
							if i = iBibSel then
								Response.Write "<option value="&i&" selected>"&Servidores.ServList.Item(i-1).Nome&"</option>"
							else
								Response.Write "<option value="&i&">"&Servidores.ServList.Item(i-1).Nome&"</option>"
							end if
						Next
					%>
					</select>
				</div>
			<% end if %>
			<div class="divForm login">
				<input type="button" class="submit" id="cadastrarUsuario" value="<%=GetTermo(global_idioma, 9298, "Cadastrar", 0) %>" onClick="return validarFormCadastroExterno();" />
			</div>
			<input type="hidden" name="origem_mobile" value="0"/>
		</form>
	</div>
	<div id="mensagem"></div>
	<script type="text/javascript">
		$(document).ready(function() {
			$("#codigo_empresa").prop("selectedIndex", -1).trigger("change");
		});
	</script>
</body>
</html>