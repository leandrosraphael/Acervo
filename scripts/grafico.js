grafico = {};

grafico._read = function (config) {
	return {
		url: config.action,
		dataType: "json",
		type: "POST"
	};
};

grafico.redimensionar = function (config) {
	var seletor = config.seletor;
	if (typeof seletor === 'string') {
		seletor = $(seletor);
	}
	$(window).on("resize", function () {
		kendo.resize(seletor);
	});
};

grafico._criarGrafico = function (config, outrasOpcoes) {
	var opcoes = $.extend(true, {
		dataSource: {
			transport: {
				read: grafico._read(config)
			},
			group: {
				field: config.groupField
			},
			sort: {
				field: config.categoryField,
				dir: "asc"
			}
		},
		theme: "material",
		title: {
			position: "top",
			text: config.titulo,
			padding: {
				bottom: 10
			}
		},
		legend: {
			visible: true,
			position: "right",
			padding: {
				top: 30
			}
		},
		chartArea: {
			background: "",
			height: 400
		},
		tooltip: {
			visible: true,
			format: config.tooltipFormat || "{0}"
		}
	}, outrasOpcoes);

	if (config.tooltipTemplate) {
		opcoes.tooltip.template = config.tooltipTemplate;
	}

	if (config.legendTemplate) {
		opcoes = $.extend(true, opcoes, {
			legend: {
				labels: {
					template: config.legendTemplate
				}
			}
		});
	}

	$(config.seletor).kendoChart(opcoes);

	grafico.redimensionar(config);
};

function labelReplace(label) {
	return label.replace("\ne\n", " e\n")
		.replace("\nde\n", " de\n");
}

function labelTemplate(e) {
	if (typeof e.value === "string") {
		return labelReplace(e.value.split(" ").join("\n"));
	}

	return e.value;
}

grafico.criarGraficoLinha = function (config) {
	/**
	* Criar gráfico linhas
	*
	* @param {Object} config - {
			seletor,
			action,
			titulo,
			groupField,
			categoryField,
			tooltipTemplate,
			tooltipFormat,
			legendTemplate,
			step
		}
	*/

	var opcoes = {
		series: [{
			type: "line",
			style: "smooth",
			field: config.dataField,
			categoryField: config.categoryField,
			name: "#= group.value #"
		}],
		categoryAxis: {
			line: {
				visible: true
			},
			labels: {
				rotation: 0,
				step: config.step
			}
		}
	};

	grafico._criarGrafico(config, opcoes);
};

grafico.criarGraficoColuna = function (config) {
	/**
	* Criar gráfico colunas
	*
	* @param {Object} config - {
			seletor,
			action,
			titulo,
			groupField
			seriesLabelTemplate,
			categoryField,
			axisDefaultsFormat,
			axisDefaultsTemplate,
			tooltipTemplate,
			tooltipFormat,
			legendTemplate,
			step
		}
	*/

	var opcoes = {
		series: [{
			type: "column",
			style: "smooth",
			field: config.dataField,
			categoryField: config.categoryField,
			name: "#= group.value #"
		}],
		categoryAxis: {
			field: config.categoryField,
			line: {
				visible: true
			},
			labels: {
				rotation: 0,
				step: config.step,
				template: labelTemplate
			}
		}
	};

	if (config.axisDefaultsFormat || config.axisDefaultsTemplate) {
		opcoes.axisDefaults = {
			labels: {}
		};

		if (config.axisDefaultsFormat) {
			opcoes.axisDefaults.labels.format = config.axisDefaultsFormat;
		}

		if (config.axisDefaultsTemplate) {
			opcoes.axisDefaults.labels.template = config.axisDefaultsTemplate;
		}
	}

	if (config.seriesLabelTemplate) {
		opcoes.seriesDefaults.labels = {
			visible: true,
			background: "transparent",
			template: config.seriesLabelTemplate
		};
	}

	grafico._criarGrafico(config, opcoes);
};