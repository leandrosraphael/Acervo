grafico.TrabalhoGraduacao = {};

grafico.TrabalhoGraduacao.criarGraficoAnoCurso = function (seletor) {

	var titulo = getTermo(global_frame.iIdioma, 9350, 'Trabalhos de Graduação', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 67, 'Ano', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 136, 'Curso', 0);

	grafico.criarGraficoLinha({
		seletor: seletor,
		action: 'asp/ajx_buscar_dados_indicador.asp?indicador=curso',
		titulo: titulo,
		groupField: 'curso',
		dataField: 'qtde',
		categoryField: 'ano',
		step: 2
	});
};

grafico.TrabalhoGraduacao.criarGraficoCursoAno = function (seletor) {

	var titulo = getTermo(global_frame.iIdioma, 9350, 'Trabalhos de Graduação', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 136, 'Curso', 0) + ' x ' +
		getTermo(global_frame.iIdioma, 67, 'Ano', 0);

	grafico.criarGraficoColuna({
		seletor: seletor,
		action: 'asp/ajx_buscar_dados_indicador.asp?indicador=curso',
		titulo: titulo,
		groupField: 'ano',
		dataField: 'qtde',
		categoryField: 'curso',
		step: 1
	});
};