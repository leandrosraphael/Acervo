var configurarPainelGeral = function () {
    grafico.Geral.criarGraficoMaterialAnoLinha('#grafico-material-ano-linha');

    grafico.Geral.criarGraficoMaterialAnoBarra('#grafico-material-ano-barra');
};

var configurarPainelTesesDissertacao = function ()
{
	grafico.TeseDissertacao.criarGraficoAnoPrograma('#grafico-tese-dissertacao-ano-programa');

	grafico.TeseDissertacao.criarGraficoProgramaAno('#grafico-tese-dissertacao-programa-ano');

	grafico.TeseDissertacao.criarGraficoMaterialAno('#grafico-tese-dissertacao-nivel-ano');
};

var configurarPainelTrabalhoGraduacao = function () {

	grafico.TrabalhoGraduacao.criarGraficoAnoCurso('#grafico-trabalho-graduacao-ano-curso');

	grafico.TrabalhoGraduacao.criarGraficoCursoAno('#grafico-trabalho-graduacao-curso-ano');
};

var configurarGradeAcessoEmpresa = function () {

	datatable.AcessoEmpresa.criarGradeAcessoEmpresa('#datatable-acesso-empresa');
};

$(function () {
	$('#indicadores').kendoPanelBar().data("kendoPanelBar").expand($('[id^="painel"]'));

    configurarPainelGeral();
	configurarPainelTesesDissertacao();
	configurarPainelTrabalhoGraduacao();
	configurarGradeAcessoEmpresa();

	setTimeout(function () {
		$.ajax({
			type: 'GET',
			url: 'asp/ajx_buscar_dados_indicador.asp?indicador=reset',
			done: function (data) { }
		})
	}, 800);
});