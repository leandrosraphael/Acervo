// ******************************************************************************
// Rotinas para as janela Modais
// ******************************************************************************

var customFechaPopup = null;
var _alturaJanela = null;
var _alturaFrame = null;

function setMascaraModal() {
	$('#ModalBG').css("position", "absolute");
	$('#ModalBG').css("top", "0px");
	$('#ModalBG').css("left", "0px");
	$('#ModalBG').css("width", "100%");
	$('#ModalBG').css("height", $(document).height());
}

function setModalPosition(wd, ht) {
	var tp = (($(window).height() - ht) / 2) + $(window).scrollTop();
	var lt = (($(window).width() - wd) / 2) + $(window).scrollLeft();

	$('#mbox').css("top", (tp < 0 ? 0 : tp) + "px");
	$('#mbox').css("left", (lt < 0 ? 0 : lt) + "px");
}

function scrollFix() {
	var wd = $('#mbox').width();
	var ht = $('#mbox').height();

	setModalPosition(wd, ht);
}

function sizeFix() {
	setMascaraModal();

	var wd = $('#mbox').width();
	var ht = $('#mbox').height();

	var winWd = $(window).width();
	var winHt = $(window).height();

	if (winWd < wd) {
		var newWd = (winWd - 10);
		if (newWd < 200) {
			newWd = 200;
		}
		$('#mbox').css("width", newWd + "px");
		$('#mbox > div').css("width", newWd + "px");
	}
	if (winHt < ht) {
		var newHt = (winHt - 10);
		if (newHt < 200) {
			newHt = 200;
		}
		$('#mbox').css("height", newHt + "px");
		$('#mbox > div').css("height", newHt + "px");
		$('#div_area_iframe').css("height", (winHt - 44) + "px");
	}

	setModalPosition(wd, ht);
}

function disable_select() {
	$('select').attr('disabled', 'disabled');
}

function enable_select() {
	$('select').removeAttr('disabled', 'disabled');
}

function sm(HTML, wd, ht) {
	setMascaraModal();

	$('#mbox').css("width", wd);
	$('#mbox').css("height", ht);

	setModalPosition(wd, ht);

	$('#mbox').html(HTML);

	$('#ModalBG').show();
	$('#mbox').show();

	//*********************************************
	// Desabilita os combos na tela
	//*********************************************
	var is_ie6 = document.all && (navigator.userAgent.toLowerCase().indexOf("msie 6.") != -1);
	if (is_ie6) {
		disable_select();
	}

	return false;
}

function hm() {
	$('#mbox').hide();
	$('#ModalBG').hide();

	//*********************************************
	// Habilita os combos na tela
	//*********************************************
	var is_ie6 = document.all && (navigator.userAgent.toLowerCase().indexOf("msie 6.") != -1);
	if (is_ie6) {
		enable_select();
	}
}

function initmb() {
	$('<div id="ModalBG"></div>').appendTo('body');
	$('<div id="mbox"></div>').appendTo('body');

	$('#mbox').hide();
	$('#ModalBG').hide();

	$(window).scroll(scrollFix);
	$(window).resize(sizeFix);
}

// ******************************************************************************

function alterarURL(HTML, URL) {
	//*********************************************************
	// Altera a URL do iFrame do Popup
	//*********************************************************

	var HTML_Final = "";
	var HTML_Parcial = "";

	var pos_iframe = HTML.indexOf("<iframe");

	HTML_Final = HTML.substring(0, pos_iframe);
	HTML_Parcial = HTML.substring(pos_iframe, HTML.length);

	var pos_src = HTML_Parcial.indexOf("src=");

	HTML_Final += HTML_Parcial.substring(0, pos_src + 5);
	HTML_Parcial = HTML_Parcial.substring(pos_src + 5, HTML_Parcial.length);

	var pos_aspas = HTML_Parcial.indexOf("'");

	HTML_Final += URL;
	HTML_Final += HTML_Parcial.substring(pos_aspas, HTML_Parcial.length);

	return HTML_Final;
}

function adicionarURL(HTML, PARAM_URL) {
	//*********************************************************
	// Adiciona um valor no final da URL do iFrame do Popup
	//*********************************************************

	var HTML_Final = "";
	var HTML_Parcial = "";

	var pos_iframe = HTML.indexOf("<iframe");

	HTML_Final = HTML.substring(0, pos_iframe);
	HTML_Parcial = HTML.substring(pos_iframe, HTML.length);

	var pos_src = HTML_Parcial.indexOf("src=");

	HTML_Final += HTML_Parcial.substring(0, pos_src + 5);
	HTML_Parcial = HTML_Parcial.substring(pos_src + 5, HTML_Parcial.length);

	var pos_aspas = HTML_Parcial.indexOf("'");
	var URL = HTML_Parcial.substring(0, pos_aspas);

	var pos_igual = PARAM_URL.indexOf("=");
	var param = PARAM_URL.substring(1, pos_igual);

	// Verifica se o parâmetro já existe
	var pos_param = URL.indexOf(param);
	if (pos_param >= 0) {
		var URL_PARCIAL = URL.substring(pos_param, URL.length);
		URL = URL.substring(0, pos_param - 1);

		var pos_nextParam = URL_PARCIAL.indexOf("&");
		if (URL_PARCIAL.indexOf("&") >= 0) {
			URL += URL_PARCIAL.substring(pos_nextParam, URL_PARCIAL.length);
		}
	}

	HTML_Final += URL;
	if (URL.indexOf(PARAM_URL) < 0) {
		HTML_Final += PARAM_URL;
	}
	HTML_Final += HTML_Parcial.substring(pos_aspas, HTML_Parcial.length);

	return HTML_Final;
}

// ******************************************************************************

function montaHTML(url, titulo, w, h, habilita_scroll, habilita_fechar, id, veio_de) {
	if (parent.hiddenFrame != null) {
		var global_frame = parent.hiddenFrame;
	} else if (parent.parent.hiddenFrame != null) {
		var global_frame = parent.parent.hiddenFrame;
	} else {
		var global_frame = parent.parent.parent.hiddenFrame;
	}

	var HTML_popup = "";
	var tipo_scroll = "";
	var func_fechar = "";
	var w_popup = 0;
	var h_popup = h + 26;

	if (habilita_scroll == true) {
		tipo_scroll = 'yes';
		w_popup = w + 8;
	} else {
		tipo_scroll = 'no';
		w_popup = w;
	}

	if (id == 2) {
		func_fechar = 'fechaPopup2';
	} else if (id == 3) {
		func_fechar = 'fechaPopup3';
	} else {
		func_fechar = 'fechaPopup';
	}

	var msg_aguarde = getTermo(global_frame.iIdioma, 32, "Carregando", 0) + "...";
	var msg_fechar = getTermo(global_frame.iIdioma, 220, "fechar", 2);

	var path = "";
	if ((veio_de != undefined) && (veio_de != "")) {
		path = "../";
	}

	HTML_popup = "<div style='width: " + w_popup + "px; height: " + h_popup + "px'>" +
		"<div class='popup_cabecalho_background popup_cabecalho_fonte' style='height: 31px'>" +
		"<span style='float: left'>" + titulo + "</span>";

	if (habilita_fechar == true) {
		HTML_popup = HTML_popup +
			"<span style='float: right'>" +
			"<a class='link_topo' href='javascript:" + func_fechar + "();' title='" + msg_fechar + "'><span class='span_imagem icon_16 icon-close' style='vertical-align:text-bottom;'></span></a>&nbsp;&nbsp;" +
			"</span>";
	}

	HTML_popup = HTML_popup +
		"</div>" +
		"<div id='div_area_iframe' class='popup_cabecalho_background' style='height: " + (h_popup - 31) + "px; padding-top: 0'>" +
		"<div id='div_loadingPop' name='div_loadingPop' style='position: absolute; top: 60px; width: " + w_popup + "px;'>" +
		"<p style='text-align: center'><span class='span_imagem icon_16 mozilla_blu'></span><br /><br />" + msg_aguarde + "</p></div>" +
		"<iframe id='popup-frame' frameborder='0' height='100%' width='100%' src='" + url + "' scrolling='" + tipo_scroll + "'></iframe>" +
		"</div></div>";

	//******************************************************************************************
	// Adiciona o HTML abaixo no IE6 para que o combo não fique por cima do div
	//******************************************************************************************
	var is_ie6 = document.all && (navigator.userAgent.toLowerCase().indexOf("msie 6.") != -1);
	if (is_ie6) {
		HTML_popup += "<iframe style=\"position: absolute; display: block; " +
			"z-index: -1; width: 100%; height: 100%; top: 0; left: 0;" +
			"filter: mask(); background-color: #ffffff;\"></iframe>";
	}

	return HTML_popup;
}

function fechaPopup(veio_de, idioma, edsPagina) {
	if (customFechaPopup) {
		var fecharCustomizado = customFechaPopup;
		customFechaPopup = null;
		fecharCustomizado();
	} else {
		// caso existir o iframe do EDS (PDF), remover
		var divAreaIframe = $('#div_area_iframe');
		if (divAreaIframe.length > 0) {
			var iFrame = divAreaIframe.find("iframe");
			if (iFrame.attr('src').indexOf('eds_full_text.') >= 0 &&
				iFrame.attr('src').indexOf('type=1') >= 0) {
				iFrame.remove();
			}
		}

		if (parent.hiddenFrame != null) {
			var global_frame = parent.hiddenFrame;
			var main_frame = parent.mainFrame;
		} else if (parent.parent.hiddenFrame != null) {
			var global_frame = parent.parent.hiddenFrame;
			var main_frame = parent.parent.mainFrame;
		} else {
			var global_frame = parent.parent.parent.hiddenFrame;
			var main_frame = parent.parent.parent.mainFrame;
		}

		hm();
		global_frame.popup_html = "";
		global_frame.popup_w = 0;
		global_frame.popup_h = 0;

		if (veio_de == "eds_full_text") {
			LinkPaginacao(global_frame.modo_busca, 'eds', edsPagina, edsPagina, 1, global_frame.modo_busca_bak, 'resultado');
		} else {
			if (global_frame.popup_refresh == true) {
				if ((global_frame.content == "resultado") || (global_frame.content == "busca_link")) {
					LinkPesquisa(global_frame.modo_busca, veio_de);
				} else if (global_frame.content == "selecao") {
					LinkSelecao(global_frame.modo_busca, veio_de);
				} else if (global_frame.content == "detalhe") {
					main_frame.location = main_frame.location + "&refresh_popup=1";
				} else {
					var pos = main_frame.location.toString().indexOf("?");
					if (pos < 0) {
						main_frame.location = main_frame.location + "?refresh_popup=1";
					} else {
						main_frame.location = main_frame.location + "&refresh_popup=1";
					}
				}
				global_frame.popup_refresh = false;
			}
		}

		document.body.style.overflow = "";
	}
}

function abrePopup(url, titulo, w, h, habilita_scroll, habilita_fechar, veio_de, auto_size, max_width) {
	if (auto_size) {
		var altura = 450;
		var largura = 680;

		var alturaJanela = $(window).height();
		var larguraJanela = $(window).width();
		altura = alturaJanela - 150;
		largura = larguraJanela - 150;

		if (max_width) {
			if (largura > max_width) {
				largura = max_width;
			}
		}
		w = largura;
		h = altura;
	}

	if (parent.hiddenFrame != null) {
		var global_frame = parent.hiddenFrame;
	} else if (parent.parent.hiddenFrame != null) {
		var global_frame = parent.parent.hiddenFrame;
	} else {
		var global_frame = parent.parent.parent.hiddenFrame;
	}

	var HTML_popup = "";
	var w_popup = 0;
	var h_popup = h + 26;

	if (habilita_scroll == true) {
		w_popup = w + 8;
	} else {
		w_popup = w;
	}

	HTML_popup = montaHTML(url, titulo, w, h, habilita_scroll, habilita_fechar, 1, veio_de);

	sm(HTML_popup, w_popup, h_popup);

	global_frame.popup_html = HTML_popup;
	global_frame.popup_w = w_popup;
	global_frame.popup_h = h_popup;
	document.body.style.overflow = "hidden";

}

function fechaPopup2() {
	if (parent.hiddenFrame != null) {
		var global_frame = parent.hiddenFrame;
	} else if (parent.parent.hiddenFrame != null) {
		var global_frame = parent.parent.hiddenFrame;
	} else {
		var global_frame = parent.parent.parent.hiddenFrame;
	}

	global_frame.popup_html2 = "";
	global_frame.popup_w2 = 0;
	global_frame.popup_h2 = 0;

	var HTML = global_frame.popup_html;
	var w = global_frame.popup_w;
	var h = global_frame.popup_h;

	sm(HTML, w, h);
}

function abrePopup2(url, titulo, w, h, habilita_scroll, habilita_fechar, veio_de) {
	if (parent.hiddenFrame != null) {
		var global_frame = parent.hiddenFrame;
	} else if (parent.parent.hiddenFrame != null) {
		var global_frame = parent.parent.hiddenFrame;
	} else {
		var global_frame = parent.parent.parent.hiddenFrame;
	}

	var HTML_popup = "";
	var w_popup = 0;
	var h_popup = h + 26;

	if (habilita_scroll == true) {
		w_popup = w + 8;
	} else {
		w_popup = w;
	}

	HTML_popup = montaHTML(url, titulo, w, h, habilita_scroll, habilita_fechar, 2, veio_de);

	sm(HTML_popup, w_popup, h_popup);

	global_frame.popup_html2 = HTML_popup;
	global_frame.popup_w2 = w_popup;
	global_frame.popup_h2 = h_popup;
}

function fechaPopup3() {
	if (parent.hiddenFrame != null) {
		var global_frame = parent.hiddenFrame;
	} else if (parent.parent.hiddenFrame != null) {
		var global_frame = parent.parent.hiddenFrame;
	} else {
		var global_frame = parent.parent.parent.hiddenFrame;
	}

	var HTML = global_frame.popup_html2;
	var w = global_frame.popup_w2;
	var h = global_frame.popup_h2;

	sm(HTML, w, h);
}

function abrePopup3(url, titulo, w, h, habilita_scroll, habilita_fechar, veio_de) {
	var HTML_popup = "";
	var w_popup = 0;
	var h_popup = h + 26;

	if (habilita_scroll == true) {
		w_popup = w + 8;
	} else {
		w_popup = w;
	}

	HTML_popup = montaHTML(url, titulo, w, h, habilita_scroll, habilita_fechar, 3, veio_de);

	sm(HTML_popup, w_popup, h_popup);
}

function fechaLoadingPopup() {
	$('#div_loadingPop').html('');
	$('#div_loadingPop').css('display', 'none');
	inserirRybena();
}

function exibeLoadingPopup() {
	if (parent.hiddenFrame != null) {
		var global_frame = parent.hiddenFrame;
	} else if (parent.parent.hiddenFrame != null) {
		var global_frame = parent.parent.hiddenFrame;
	} else {
		var global_frame = parent.parent.parent.hiddenFrame;
	}
	var div = $('#div_loadingPop');
	if (div.length) {
		var msg_aguarde = getTermo(global_frame.iIdioma, 32, "Carregando", 0) + "...";
		div.html("<p style='text-align: center'><span class='span_imagem icon_16 mozilla_blu'></span><br /><br />" + msg_aguarde + "</p>");
	}
}

function inserirRybena() {
	if (global_integracao_rybena == 1) {
		var addScript = function (src) {
			var $context = $('iframe#popup-frame');
			var script = $context.createElement('script');
			script.type = 'text/javascript';
			script.src = src;
			$context.contents().find('head')[0].appendChild(script);
		}

		addScript('../scripts/plugin-iframe.js');
	}
}

function alterarAlturaJanela(tamanho) {
	$("div#mbox").css({ height: tamanho + "px" });
	$("div#div_area_iframe").css({ height: tamanho + "px" });
}

function salvarAlturaJanela() {
	_alturaJanela = parseInt($("div#mbox").css("height").replace("px", ""), 10)
	_alturaFrame = parseInt($("div#div_area_iframe").css("height").replace("px", ""), 10);
}

function restaurarAlturaJanela() {
	if (_alturaJanela && _alturaFrame) {
		$("div#mbox").css({ height: _alturaJanela + "px" });
		$("div#div_area_iframe").css({ height: _alturaFrame + "px" });
	}
	return true;
}