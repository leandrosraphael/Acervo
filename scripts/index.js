var global_frame = null;
if (parent.hiddenFrame != null) {
    global_frame = parent.hiddenFrame;
}

function exibeBuscaAvancada() {
    var texto = getTermo(global_frame.iIdioma, 6896, 'Ocultar busca avançada', 0);

    $('#buscaAvancadaLegLink').text(texto);
    $('#div_leg_avancada').show();

    $('#buscaAvancadaLegLink').unbind('click');
    $('#buscaAvancadaLegLink').click(ocultarBuscaAvancada);

    if (global_frame != null) {
        global_frame.buscaAvancadaLegislacao = true;
    }
}

function ocultarBuscaAvancada() {
    var texto = getTermo(global_frame.iIdioma, 6895, 'Exibir busca avançada', 0);

    $('#buscaAvancadaLegLink').text(texto);
    $('#div_leg_avancada').hide();

    $('#buscaAvancadaLegLink').unbind('click');
    $('#buscaAvancadaLegLink').click(exibeBuscaAvancada);

    if (global_frame != null) {
        global_frame.buscaAvancadaLegislacao = false;

        ResetarLegAvancada();
        atribuiLegs();
    }
}

function BuscarFaceta(modo, servidor, faceta) {
	var textoCarregando = getTermo(global_frame.iIdioma, 32, 'Carregando', 0) + '...';
	var textoFiltro = getTermo(global_frame.iIdioma, 1818, 'Filtros', 0);

    var htmlFacetas =
        '<table class="tab-busca-facetada max_width remover_bordas_padding centro">' +
        '<tr style="height: 26px"><td class="td_faceta_cabecalho">&nbsp;' + textoFiltro + '</td></tr>' +
        '<tr><td class="td_facetas">' +
        '<span class="span_imagem icon_16 mozilla_blu"></span><br /><br />' + textoCarregando +
        '</td></tr>' +
        '</table>';

    $('#facet_aba_' + servidor).html(htmlFacetas);

    var htmlResultado = '<br /><span class="span_imagem icon_16 mozilla_blu"><br /><br />' + textoCarregando;

    $('#p_div_aba' + servidor + '_resultado').html(htmlResultado);

    SubmeteBuscaFrame(modo, servidor, faceta);
}

function AplicarFacetaVerMais(servidor, idFaceta, facetaVerMais) {
    var form = $('#faceta-form-srv' + servidor);

    var facetaOriginal = $(form).serialize();
    var faceta = '';

    if (facetaOriginal != '') {
        while (facetaOriginal.indexOf('=') > 0) {
            var chave = facetaOriginal.substring(0, facetaOriginal.indexOf('='));
            if (chave != idFaceta) {
                if (faceta != '') {
                    faceta += '&';
                }
                if (facetaOriginal.indexOf('&') == -1) {
                    faceta += facetaOriginal;
                } else {
                    faceta += facetaOriginal.substring(0, facetaOriginal.indexOf('&'));
                }
            }

            if (facetaOriginal.indexOf('&') == -1) {
                facetaOriginal = '';
            } else {
                facetaOriginal = facetaOriginal.substring(facetaOriginal.indexOf('&') + 1);
            }
        }
    }

    if (facetaVerMais != '') {
        if (faceta != '') {
            faceta += '&';
        }
        faceta += facetaVerMais;
    }

    if (faceta != '') {
        faceta += '&';
    }

    faceta += 'alterou=1';
        
    BuscarFaceta(1, servidor, faceta);
}

function AplicarFacetaVerMaisLeg(servidor, idFaceta, facetaVerMais) {
    var form = $('#faceta-form-srv' + servidor);

    var facetaOriginal = $(form).serialize();
    var faceta = '';

    if (facetaOriginal != '') {
        while (facetaOriginal.indexOf('=') > 0) {
            var chave = facetaOriginal.substring(0, facetaOriginal.indexOf('='));
            if (chave != idFaceta) {
                if (faceta != '') {
                    faceta += '&';
                }
                if (facetaOriginal.indexOf('&') == -1) {
                    faceta += facetaOriginal;
                } else {
                    faceta += facetaOriginal.substring(0, facetaOriginal.indexOf('&'));
                }
            }

            if (facetaOriginal.indexOf('&') == -1) {
                facetaOriginal = '';
            } else {
                facetaOriginal = facetaOriginal.substring(facetaOriginal.indexOf('&') + 1);
            }
        }
    }

    if (facetaVerMais != '') {
        if (faceta != '') {
            faceta += '&';
        }
        faceta += facetaVerMais;
    }

    if (faceta != '') {
        faceta += '&';
    }

    faceta += 'alterou=1';

    BuscarFaceta(5, servidor, faceta);
}

function configurarBuscaFacetada() {
    $('body').on('click', 'li.item-faceta > input[type="checkbox"]', function () {
        var that = this;
        setTimeout(function () {
            var form = $(that).closest('form');
            var servidor = $(form).attr('data-servidor');
            var faceta = form.serialize();

            if (faceta != '') {
                faceta += '&';
            }

            faceta += 'alterou=1';

            BuscarFaceta(1, servidor, faceta);
        }, 10);
    });

    $('body').on('click', 'li.item-faceta-leg > input[type="checkbox"]', function () {
        var that = this;
        setTimeout(function () {
            var form = $(that).closest('form');
            var servidor = $(form).attr('data-servidor');
            var faceta = form.serialize();

            if (faceta != '') {
                faceta += '&';
            }

            faceta += 'alterou=1';

            BuscarFaceta(5, servidor, faceta);
        }, 10);
    });

    $('body').on('click', '.remover-faceta', function (e) {
        var checkId = $(this).attr('data-checkid');
        $('#' + checkId).click();
        e.preventDefault();
    });

    $('body').on('click', '.titulo-faceta', function (e) {
        var expandido = $(this).hasClass('expandida');
        var faceta = $(this).closest('ul');
        var contador = 1;

        var div = $(this).find('div');

        if (expandido) {
            $(div).removeClass('icon-small-down-b');
            $(div).addClass('icon-small-next-b');
        } else {
            $(div).removeClass('icon-small-next-b');
            $(div).addClass('icon-small-down-b');
        }

        $(faceta).find('li').each(function () {
            if (contador > 1) {
                if (expandido) {
                    $(this).css('display', 'none');
                } else {
                    $(this).css('display', 'block');
                }
            } else {
                if (expandido) {
                    $(this).removeClass('expandida');
                } else {
                    $(this).addClass('expandida');
                }
            }
            contador++;
        });
    });

    $('body').on('click', '.item-faceta.ver-mais', function () {
        var form = $(this).closest('form');
        var ul = $(this).closest('ul');
        var servidor = $(form).attr('data-servidor');
        var indiceFaceta = $(ul).attr('data-indice');

        var mensagem = getTermo(global_frame.iIdioma, 1818, 'Filtros', 0);
        var url = 'asp/exibirFacetaCompleta.asp?iIndexSrv=' + servidor + '&iFaceta=' + indiceFaceta + getGlobalUrlParams();
        abrePopup(url, mensagem, 380, 470, false, false);
    });

    $('body').on('click', '.item-faceta-leg.ver-mais', function () {
        var form = $(this).closest('form');
        var ul = $(this).closest('ul');
        var servidor = $(form).attr('data-servidor');
        var indiceFaceta = $(ul).attr('data-indice');

        var mensagem = getTermo(global_frame.iIdioma, 1818, 'Filtros', 0);
        var url = 'asp/exibirFacetaCompletaLeg.asp?iIndexSrv=' + servidor + '&iFaceta=' + indiceFaceta + getGlobalUrlParams();
        abrePopup(url, mensagem, 380, 470, false, false);
    });
}

function LinkContraste() {
    $('html').toggleClass('contraste');

    if ($('html').hasClass('contraste')) {
        $.cookie('contraste', '1', { expires: 2 });
        $('#home_slide_content_div').addClass('slide_mensagens_topo');
        $('#slide_mensagens_div').addClass('slide_mensagens_topo');
        $('#slide_mensagens_handle_fundo_div').addClass('slide_mensagens_topo');
        $('#slide_mensagens_handle').addClass('slide_mensagens_topo');
    } else {
        $.cookie('contraste', '0', { expires: 2 });
        $('#home_slide_content_div').removeClass('slide_mensagens_topo');
        $('#slide_mensagens_div').removeClass('slide_mensagens_topo');
        $('#slide_mensagens_handle_fundo_div').removeClass('slide_mensagens_topo');
        $('#slide_mensagens_handle').removeClass('slide_mensagens_topo');
    }
}

function atualizaselecaoUltimasAquisicoes(combo) {
    var indexSrv = $(combo).val();
	atualizaUltimasAquisicoes(indexSrv);
}

function atualizaUltimasAquisicoes(indexSrv) {
	$.ajax({
		type: 'POST',
		url: ext + '/ajxUltimasAquisicoes.' + ext + '?iIndexSrv=' + indexSrv + getGlobalUrlParams(),
		data: "",
		success: function (data) {
			$('#ultimas_aquisicoes_area_conteudo_div').html(data);
			if ($.trim(data) != '') {
				sessionStorage.iIndexSrvAquisicao = indexSrv;
			}
		}
	});
}

function atualiza_Lista_Favoritos(indexSrv, listaSelecionada) {
	$.ajax({
		type: 'POST',
		url: ext + '/ajxListaFavoritos.' + ext + '?iIndexSrv=' + indexSrv + '&listaSelecionada=' + listaSelecionada + getGlobalUrlParams(),
		data: "",
		success: function (data) {
			$('#combo_lista_favoritos').html(data);
			if ($.trim(data) != '') {
				visualizar_favoritos(indexSrv);
				AtualizaListaFavoritos();
			}
		}
	});
}

function carrega_titulos_Favoritos(indexSrv, listaSelecionada) {
	$.ajax({
		type: 'POST',
		url: ext + '/ajxVisualizarListaFavoritos.' + ext + '?iIndexSrv=' + indexSrv + '&listaSelecionada=' + listaSelecionada + getGlobalUrlParams(),
		data: "",
		success: function (data) {
			$('#ficha_favoritos').html(data);
			if ($.trim(data) != '') {
			}
		}
	});
}

$(document).ready(function () {
    if (global_frame != null && global_frame.buscaAvancadaLegislacao != undefined) {
        if (global_frame.buscaAvancadaLegislacao == true) {
            exibeBuscaAvancada();
        } else {
            ocultarBuscaAvancada();
        }
    } else {
        ocultarBuscaAvancada();
    }

    initmb(); 

    criaLink();

    configurarBuscaFacetada();
});

function obterLevantamentosBibliografico(Servidor) {
    $.ajax({
        type: 'POST',
        url: ext + '/ajxLevantamentoBib.' + ext + '?iIndexSrv=' + Servidor + getGlobalUrlParams(),
        data: "",
        success: function (data) {
            $('#levantamento_bibibliografico').html(data);
        }
    });
}

function criaLink() {
	//Cria um link logo abaixo do banner (Adequação metodista - BBLA)
	if (parent.hiddenFrame != undefined) {
		if ((parent.hiddenFrame.iBusca_Projeto > 0) && (global_numero_serie == 3386)) {
			var oTable = document.getElementById('tab_principal');

			var oRow1 = oTable.insertRow(1);

			var aRows = oTable.rows;

			var aCells = oRow1.cells;

			var oCell1_1 = aRows(oRow1.rowIndex).insertCell(aCells.length);

			oCell1_1.innerHTML = "&nbsp;&nbsp;<a class='link_menu1' href='http://www.metodista.br/biblica' target='_top'><img src='imagens/casa.gif' alt='' />&nbsp;Voltar à Página Inicial</a>&nbsp;";
			oCell1_1.bgColor = "<%=css_menu_background%>";

			oCell1_1.colSpan = 2;
			oCell1_1.height = 20;
			oCell1_1.align = "left";
		}
	}
}


function abreLogin(destino, tituloDestino, parametrosExtras, habilitarRolagem, habilitarFechar) {
	var tituloJanela = tituloDestino || getTermo(global_frame.iIdioma, 6628, 'Entrar', 0);
	var url = ext + '/login.' + ext;
	var alturaJanela, larguraJanela, modo_busca;

	if (global_frame) {
		alturaJanela = global_frame.alturaJanelaLogin;
		modo_busca = global_frame.modo_busca;
	} else {
		alturaJanela = 320;
		modo_busca = 'rapida';
	}

	larguraJanela = 320;

	url = url + '?modo_busca=' + modo_busca;
	url = url + '&veio_de=' + destino;

	if (parametrosExtras && parametrosExtras != "") {
		url = url + parametrosExtras;
	}
	url = url + getGlobalUrlParams();

	if (destino == 'reserva') {
		alturaJanela = 400;
		larguraJanela = 380;
	} else if (destino == 'favoritos') {
		alturaJanela = 320;
	} else if (destino == 'avaliacao_votar') {
		alturaJanela = 250;
	}
	setTimeout(function () {
		abrePopup(url, tituloJanela, larguraJanela, alturaJanela, habilitarRolagem, habilitarFechar);
	},50);
}