function BuscarTotalizadoresTeseDissertacao() {
	var anoInicial = document.getElementById("ano_inicial_filtro").value;
	var anoFinal = document.getElementById("ano_final_filtro").value;
	var tipoFiltro = document.getElementById("tipo_filtro").value;
	var material = document.getElementById("filtro_material").value;

	if (FiltroAnoValido(anoInicial, anoFinal, tipoFiltro)) {
		$.ajax({
			type: 'GET',
			url: 'asp/ajxBuscarDadosPesquisaTeseDissertacao.asp',
			data: {
				"anoInicial": anoInicial,
				"anoFinal": anoFinal,
				"tipoFiltro": tipoFiltro,
				"material": material,
				"idioma": global_frame.iIdioma
			},
			beforeSend: function () {
				$("div#carregando_consulta").removeClass("esconder-loading").addClass("exibir-loading");
			},
			success: function (data) {
				divResultado = $('#dados-tese-dissertacao').html(data);
			},
			complete: function () {
				$("div#carregando_consulta").removeClass("exibir-loading").addClass("esconder-loading");
			}
		});
	}
}

function LimparBuscaTeseDissertacao() {
	$("#filtro_material").val(0);
	$("#filtro_material").multiselect().multiselect("refresh");
	$("#filtro_material").change();

	$("#tipo_filtro").val(0);
	$("#tipo_filtro").multiselect().multiselect("refresh");
	$("#tipo_filtro").change();

	$("#ano_inicial_filtro").val("");
	$("#ano_final_filtro").val("");
}

$(function () {
	$("#btnBuscarTeseDissertacao").click(BuscarTotalizadoresTeseDissertacao);
	$("#btnBuscarTeseDissertacao").click();

	$("#btnLimparFiltroTeseDissertacao").click(LimparBuscaTeseDissertacao);
});
