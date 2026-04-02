datatable.AcessoEmpresa = {};

datatable._ObterConfiguracaoLinguagem = function () {

	return {
		emptyTable: getTermo(global_frame.iIdioma, 9353, "Não há nenhum registro de acesso.", 0),
		info: "",
		infoEmpty: "",
		infoPostFix: "",
		infoThousands: ".",
		lengthMenu: getTermo(global_frame.iIdioma, 827, "Número de resultados por página", 0) + ": _MENU_ ",
		loadingRecords: getTermo(global_frame.iIdioma, 32, "Carregando", 0) + "...",
			processing: getTermo(global_frame.iIdioma, 6670, "Processando", 0) + "...",
		paginate: {
			next: getTermo(global_frame.iIdioma, 9354, "Próximo", 0),
			previous: getTermo(global_frame.iIdioma, 1326, "Anterior", 0)
		}
	};
};

datatable.AcessoEmpresa.criarGradeAcessoEmpresa = function (seletor) {

	datatable.criarGrade({
		seletor: seletor,
		searching: false,
		ajax: 'asp/ajx_buscar_dados_indicador.asp?indicador=empresa',
		language: datatable._ObterConfiguracaoLinguagem()
	});
};