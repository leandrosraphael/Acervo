var ext = 'asp';
function limparDadosLogin() {
	document.login.codigo.value = "";
	document.login.senha.value = "";
}

function SenhaEmBranco(modo_busca, tipoBanner, idioma) {
	document.location = '../index.' + ext + '?modo_busca=' + modo_busca + '&content=troca_senha&iBanner=' + tipoBanner + "&iIdioma=" + idioma;
}

function FocoSenha() {
	try {
		document.login.senha.focus();
	} catch (er) { }
}

function FocoLogin() {
	try {
		document.login.codigo.focus();
	} catch (er) { }
}

function SetRefreshURL() {
	if (parent.parent.hiddenFrame != null) {
		Hiddenfrm = parent.parent.hiddenFrame;
		Hiddenfrm.popup_refresh = true;
	}
}

function fecharLoginStart(Parametros) {
	var url = "../spacer." + ext + "?" + Parametros;
	document.location = url;
}

function voltarLoginStart() {
	document.location = 'login_start.asp';
}

function LinkLembrarSenhaLogin(rm) {
	var tam_codigo = document.login.codigo.value.length;
	if (document.login.servidor != null) {
		var serv = document.login.servidor.value;
		var url = 'lembrar_start.' + ext + '?codigo=' + document.login.codigo.value + '&servidor=' + serv + getGlobalUrlParams();
	} else {
		var url = 'lembrar_start.' + ext + '?codigo=' + document.login.codigo.value + getGlobalUrlParams();
	}
	if (tam_codigo == 0) {
		var mensagem = getTermo(global_frame.iIdioma, 1312, 'Para ver a lembrança de sua senha, preencha o campo %s e clique em Lembrar senha.', 0);
		var identificador = "";

		if (rm == 0) {
			identificador = getTermo(global_frame.iIdioma, 82, "Matrícula", 0);
		}
		else
			if (rm == 1) {
				identificador = getTermo(global_frame.iIdioma, 81, "Código", 0);
			}
			else {
				identificador = getTermo(global_frame.iIdioma, 95, "Login", 0);
			}
		mensagem = mensagem.replace("%s", identificador);
		alert(mensagem);
	} else {
		document.location = url;
	}
}

function SolicitarEmail(identificacao, servidor, tipoBanner, idioma) {
	document.location = "lembrar_start." + ext + "?acao=envia_email&codigo=" + identificacao + "&servidor=" + servidor + "&iBanner=" + tipoBanner + "&iIdioma=" + idioma;
}
