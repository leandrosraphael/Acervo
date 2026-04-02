function exibeCampoDescricao() {
	var codigo = $('#lista_favorito').val();

	if (codigo == 0) {
		$('#div_inputListaFavorito').css('display', 'inline');
	} else {
		$('#div_inputListaFavorito').css('display', 'none');
	}
}

function incluirFavoritos(Codigos, Servidor, Idioma) {
	var codigoLista = $('#lista_favorito').val();
	if (codigoLista == 0) {
		var descricao = $('#inputListaFavorito').val();
		if (descricao == '') {
			alert(getTermo(Idioma, 3927, 'A descrição deve ser preenchida.', 0));
		} else {
			url = "salvar_favoritos.asp?descricao=" + descricao + "&codigos=" + Codigos + "&servidor=" + Servidor + "&idioma=" + Idioma;
			document.location.href = url;
		}
	} else {
		url = "salvar_favoritos.asp?codigoLista=" + codigoLista + "&codigos=" + Codigos + "&servidor=" + Servidor + "&idioma=" + Idioma;
		document.location.href = url;
	}
}

$(document).ready(function () {
	exibeCampoDescricao();
});