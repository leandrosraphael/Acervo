function extensaoVideo(extensao) {
	if ((extensao == 'ogg') || (extensao == 'webm') || (extensao == 'mp4')) {
		return true;
	} else {
		return false;
	}
}

function abrirPlayerHtml5(extensao) {
	if (extensaoVideo(extensao)) {
		if (Modernizr.video) {
			if ((extensao == 'ogg') && (Modernizr.video.ogg)) {
				return true;
			} else if ((extensao == 'webm') && (Modernizr.video.webm)) {
				return true;
			} else if ((extensao == 'mp4') && (Modernizr.video.h264)) {
				return true;
			}
		}
	}

	return false;
}

function tratarUrl(url, extensao) {
	extensao = extensao.toLowerCase();
	if (abrirPlayerHtml5(extensao)) {
		return url + '&video=1&extensao=' + extensao;
	} else {
		return url;
	}
}

function abrirAba(url, extensao) {
	var urlMidia = tratarUrl(url, extensao);
	window.open(urlMidia);
}

$(document).ready(function () {
	$('.link-repositorio').click(function (e) {
		var url = $(e.target).attr('href');
		var ext = $(e.target).attr('data-extensao');

		if (url.indexOf('download.') >= 0) {
			e.preventDefault();
			abrirAba(url, ext);
		}
	});
});