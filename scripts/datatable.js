datatable = {};

datatable.criarGrade = function (config) {

	$(config.seletor).DataTable({
		'processing': true,
		'serverSide': false,
		'searching': config.searching,
		'ajax': config.ajax,
		'language': config.language
	});
};