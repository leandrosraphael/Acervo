grafico.Geral = {};

grafico.Geral.criarGraficoMaterialAnoLinha = function (seletor) {

	var titulo = getTermo(global_frame.iIdioma, 175, 'Material', 0) + ' x ' + getTermo(global_frame.iIdioma, 67, 'Ano', 0);

	grafico.criarGraficoLinha({
		seletor: seletor,
		action: 'asp/ajx_buscar_dados_indicador.asp?indicador=geral_material',
		titulo: titulo,
		groupField: 'material',
		dataField: 'qtde',
		categoryField: 'ano',
		step: 2
	});
};

grafico.Geral.criarGraficoMaterialAnoBarra = function (seletor) {

    var titulo = getTermo(global_frame.iIdioma, 175, 'Material', 0) + ' x ' + getTermo(global_frame.iIdioma, 67, 'Ano', 0);

    grafico.criarGraficoColuna({
        seletor: seletor,
        action: 'asp/ajx_buscar_dados_indicador.asp?indicador=geral_material',
        titulo: titulo,
        groupField: 'material',
        dataField: 'qtde',
        categoryField: 'ano',
        step: 2
    });
};