function emprestar() {
	var exemplar = document.querySelector("div#frmEmprestimo input[name='exemplar']").value;
	if (exemplar) {
		var usuario = document.querySelector("div#frmEmprestimo input[name='usuario']").value;
		var idioma = document.querySelector("div#frmEmprestimo input[name='idioma']").value;
		var tipo = document.querySelector("div#frmEmprestimo input[name='tipo']").value;
		validarExemplar(usuario, exemplar, tipo, idioma);
	} else {
		alert(getTermo(global_frame.iIdioma, 442, 'É preciso selecionar um exemplar.', 0));
	}
	return false;
}

function validarExemplar(usuario, exemplar, tipo, idioma) {
	$.ajax({
		type: 'POST',
		url: "asp/validarExemplar.asp",
		data: "exemplar = " + exemplar + "&usuario=" + usuario + "&tipo=" + tipo + "&idioma=" + idioma,
		success: function (data) {
			document.querySelector("div#frmEmprestimo div[class='ficha-emprestimo']").innerHTML = data;
		},
		error: function (x, s, e) {
			var MsgErro = "Erro ao validar o exemplar.";
			MsgErro += (e ? "Erro: " + e : "");
			console.error(MsgErro);
		}
	});
}

function VoltarTelaDevolucao(modo) {
	parent.mainFrame.location = "index.asp?modo_busca=" + modo + "&content=circulacoes&acao=emprestimo" + getGlobalUrlParams();
}

function keyPressEmprestimo(ev, name) {
	if (ev.charCode == 13) {
		emprestar();
	} else {
		if (name == "codigo") {
			return ev.charCode >= 48 && ev.charCode <= 57;
		} else {
			return true;
		}
	}
};

function emprestarExemplar(queryString) {
	var urlEmprestimo = "asp/emprestarExemplar.asp?" + queryString;
	$.ajax({
		type: 'POST',
		url: urlEmprestimo,
		success: function (data) {
			document.querySelector("div#frmEmprestimo div[class='ficha-emprestimo']").innerHTML = data;
		},
		error: function (x, s, e) {
			var MsgErro = "Erro ao emprestar o exemplar.";
			MsgErro += (e ? "Erro: " + e : "");
			console.error(MsgErro);
		}
	});
	return false;
}