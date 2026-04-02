var ext = "asp";

function renomearFavorito(CodigoFavorito, Servidor, Idioma) {
	var descricao = $('#inputListaFavorito').val();
	if (descricao == '') {
		alert(getTermo(Idioma, 3927, 'A descrição deve ser preenchida.', 0));
	} else {
		url = "renomear_favoritos." + ext + "?descricao=" + descricao + "&codigoFavorito=" + CodigoFavorito + "&servidor=" + Servidor + "&idioma=" + Idioma;
		document.location.href = url;
	}
}

function altera_descricao_favoritos(srv) {
	var msg_minhasel = getTermo(global_frame.iIdioma, 8322, 'Renomear', 0);
	var lista = document.getElementById('lista_favorito');
	var codigoLista = lista.options[lista.selectedIndex].value
	var descricaoLista = lista.options[lista.selectedIndex].text;
	abrePopup(ext + "/nova_descricao_favoritos." + ext + "?codigoLista=" + codigoLista + "&descricaoLista=" + descricaoLista + "&servidor=" + srv + getGlobalUrlParams(), msg_minhasel,
		320, 200, false, false);
}

function excluir_favorito(srv, mensagem) {
	var titulo_msg = getTermo(global_frame.iIdioma, 167, 'Excluir', 0);
	var codigoLista = $('#lista_favorito').val();
	var url = ext + "/excluir_favoritos." + ext + "?codigoFavorito=" + codigoLista + "&servidor=" + srv + getGlobalUrlParams();
	abrePopup(ext + "/confirmation." + ext + "?url=" + url + "&titulo_msg=" + titulo_msg + "&mensagem=" + mensagem + "&height=320&Weight=300", titulo_msg,
			320, 200, false, false);
}

function excluir_titulo_favorito(CodigoTitulo, favorito, srv) {
	var msg_minhasel = getTermo(global_frame.iIdioma, 167, 'Excluir', 0);
	abrePopup(ext + "/excluir_favoritos." + ext + "?codigoFavorito=" + favorito + "&codigoTitulo=" + CodigoTitulo + "&servidor=" + srv + getGlobalUrlParams(), msg_minhasel,
			320, 300, false, false);
}

function visualizar_favoritos(srv) {
	$("#loading-favorito").toggleClass("loading-favorito");
	var msg_minhasel = getTermo(global_frame.iIdioma, 12, 'Visualizar', 0);
	var codigoLista = $('#lista_favorito').val();
	carrega_titulos_Favoritos(srv, codigoLista);
}

function AtualizaListaFavoritos() {
	$(document).ready(function () {
		$(".styled_combo").each(function () {
			var id = $(this).attr("id");
			$("#" + id).multiselect({
				multiple: false,
				minWidth: 20,
				selectedList: 1,
				header: false,
				open: function () {
					var sel = $("#" + id).multiselect().multiselect("getChecked");
					var selLi = sel.parent().parent();
					var selUl = selLi.parent();
					var index = selUl.find("li").index(selLi);
					$(selUl).scrollTop($(selLi).height() * index);

				}
			});
		});
	});
}
