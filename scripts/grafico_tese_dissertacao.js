grafico.TeseDissertacao = {};

grafico.TeseDissertacao.criarGraficoAnoPrograma = function (seletor) {

	var titulo = getTermo(global_frame.iIdioma, 2194, 'Teses', 0) + '/' +
		getTermo(global_frame.iIdioma, 9396, 'Dissertações', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 67, 'Ano', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 738, 'Programa', 0);

	grafico.criarGraficoLinha({
		seletor: seletor,
		action: 'asp/ajx_buscar_dados_indicador.asp?indicador=programa',
		titulo: titulo,
		groupField: 'programa',
		dataField: 'qtde',
		categoryField: 'ano',
		step: 2
	});
};

grafico.TeseDissertacao.criarGraficoProgramaAno = function (seletor) {

	var titulo = getTermo(global_frame.iIdioma, 2194, 'Teses', 0) + '/' +
		getTermo(global_frame.iIdioma, 9396, 'Dissertações', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 738, 'Programa', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 67, 'Ano', 0);

	grafico.criarGraficoColuna({
		seletor: seletor,
		action: 'asp/ajx_buscar_dados_indicador.asp?indicador=programa',
		titulo: titulo,
		groupField: 'ano',
		dataField: 'qtde',
		categoryField: 'programa',
		step: 1
	});
};

grafico.TeseDissertacao.criarGraficoMaterialAno = function (seletor) {

	var titulo = getTermo(global_frame.iIdioma, 2194, 'Teses', 0) + '/' +
		getTermo(global_frame.iIdioma, 9396, 'Dissertações', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 175, 'Material', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 67, 'Ano', 0);

	grafico.criarGraficoColuna({
		seletor: seletor,
		action: 'asp/ajx_buscar_dados_indicador.asp?indicador=material',
		titulo: titulo,
		groupField: 'material',
		dataField: 'qtde',
		categoryField: 'ano',
		step: 2
	});
};