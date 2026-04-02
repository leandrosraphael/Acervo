
function emprestar_fgv(midia, usuario) {
	var url = "ajxEmprestarMidiaDigital.asp?midia=" + midia + "&usuario=" + usuario;

	executarAjax(url);
}

function reservar_fgv(midia, usuario) {
	var url = "ajxReservarMidiaDigital.asp?acao=reservar&midia=" + midia + "&usuario=" + usuario;

	executarAjaxReserva(url); 
}

function executarAjaxReserva(url) {
	$.ajax({
		type: 'GET',
		url: url,
		success: function (texto) {
			mensagem_fgv(texto);
		},
		error: function (jqXHR, textStatus, errorThrown) {
			mensagem_fgv("Erro ao executar a operação.");
		}
	});
}

function executarAjax(url) {
	$.ajax({
		url: url,
		type: 'GET',
		statusCode: {
			200: function (response) {
				exibir_link_fgv(response);
			},
			203: function (response) {
				mensagem_fgv(response);
			},
			500: function (response) {
				mensagem_fgv(response.responseText);
			}
		}
	});
}

function mensagem_fgv(mensagem) {
	var url = ext + "/mensagem-popup." + ext + "?mensagem=" + escape(mensagem);
	parent.abrePopup(url, getTermo(global_frame.iIdioma, 9074, "Empréstimo de titulo digital", 0), 500, 390, true, true);
}

function exibir_link_fgv(link) {
	var url = ext + "/pop_exibe_link.asp?link=" + escape(link);
	parent.abrePopup(url, getTermo(global_frame.iIdioma, 9074, "Empréstimo de titulo digital", 0), 500, 390, true, true);
}