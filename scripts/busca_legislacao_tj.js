function LinkMarcarTodos(seletor) {
	$('.' + seletor).prop('checked', true);
	$('.' + seletor + '_codigos').val('');

	var array = [];
	var aChk = $('.' + seletor);
	for (var i = 0; i < aChk.length; i++) {
		array.push(aChk[i].value);
	}
	$('.' + seletor + '_codigos').val(array.toString());
	atualizaLeg_campos();
}

function LinkDesmarcarTodos(seletor) {
	$('.' + seletor).prop('checked', false);
	$('.' + seletor + '_codigos').val('');
	atualizaLeg_campos();
}

function atualizaSelecao(ckSeletor, seletor) {
	var array = [];
	if ($('.' + seletor + '_codigos').val() != '') {
		array = $('.' + seletor + '_codigos').val().split(",");
	}

	if (ckSeletor.checked) {
		array.push(ckSeletor.value);
	} else {
		for (var i = 0; i < array.length; i++) {
			if (array[i] == ckSeletor.value) {
				array.splice(i, 1);
				break;
			}
		}
	}
	$('.' + seletor + '_codigos').val(array.toString());
}

function refreshCkekBox(seletor) {
	var array = [];
	var aChk = $('.' + seletor);
	if ($('.' + seletor + '_codigos').val() != '') {
		array = $('.' + seletor + '_codigos').val().split(",");
		array.sort();
		for (var i = 0; i < aChk.length; i++) {
			for (var j = 0; j < array.length; j++) {
				if (aChk[i].value == array[j]) {
					aChk[i].checked = true;
				}
			}
		}
	}
}

function ExibirMaisNormas(quantidade) {
	var selecionados = parent.hiddenFrame.ckNorma_codigos;
	$.ajax({
		type: 'POST',
		url: ext + '/ajxNormasTj.' + ext + '?quantidade=' + quantidade + '&normasSelecionadas=' + selecionados,
		data: "",
		success: function (data) {
			$('#div_selecao_norma_tj').html(data);
			if ($.trim(data) != '') {
				refreshCkekBox('ckNorma');
			}
		}
	});
}

function ExibirMaisOrgaoOrigem(quantidade) {
	var selecionados = parent.hiddenFrame.ckOrgaoOrigem_codigos;
	$.ajax({
		type: 'POST',
		url: ext + '/ajxOrgaoOrigemTj.' + ext + '?quantidade=' + quantidade + '&orgaosSelecionados=' + selecionados,
		data: "",
		success: function (data) {
			$('#div_origem_tj').html(data);
			if ($.trim(data) != '') {
				refreshCkekBox('ckOrgaoOrigem');
			}
		}
	});
}

$(document).ready(function () {
	if (parent.hiddenFrame.modo_busca == 'legislacao') {
		ExibirMaisNormas(8);
		ExibirMaisOrgaoOrigem(8);
	}
});

