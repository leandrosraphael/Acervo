function buscarResumoTrabalhosGraduacao() {
	var anoInicial = document.getElementById("ano_inicial_filtro").value;
	var anoFinal = document.getElementById("ano_final_filtro").value;
	var tipoFiltro = document.getElementById("tipo_filtro").value;

	if (FiltroAnoValido(anoInicial, anoFinal, tipoFiltro)) {
		$.ajax({
			type: 'GET',
			url: 'asp/ajxbuscar_resumos_trabalho_graduacao.asp',
			data: {
				"anoInicial": anoInicial,
				"anoFinal": anoFinal,
				"tipoFiltro": tipoFiltro
			},
			beforeSend: function () {
				$("div#carregando_consulta").removeClass("esconder-loading").addClass("exibir-loading");
			},
			success: function (data) {
				divResultado = $('#resumos-trabalho-graduacao').html(data);
			},
			complete: function () {
				$("div#carregando_consulta").removeClass("exibir-loading").addClass("esconder-loading");
			}
		});
	}
}

function LimparBuscaTrabalhoGraduacao() {
	$("#tipo_filtro").val(0);
	$("#tipo_filtro").multiselect().multiselect("refresh");
	$("#tipo_filtro").change();

	$("#ano_inicial_filtro").val("");
	$("#ano_final_filtro").val("");
}

$(function() {
	$("#btnBuscarResumoTG").click(buscarResumoTrabalhosGraduacao);
	$("#btnBuscarResumoTG").click();

	$("#btnLimparFiltroTrabalhoGraduacao").click(LimparBuscaTrabalhoGraduacao);
});