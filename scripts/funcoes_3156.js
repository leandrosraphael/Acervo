var ext = "asp";

var global_frame = null;
var main_frame = null;
if (parent.hiddenFrame != null) {
	var global_frame = parent.hiddenFrame;
	var main_frame = parent.mainFrame;
} else if (parent.parent.hiddenFrame != null) {
	var global_frame = parent.parent.hiddenFrame;
	var main_frame = parent.parent.mainFrame;
} else {
	var global_frame = parent.parent.parent.hiddenFrame;
	var main_frame = parent.parent.parent.mainFrame;
};

function pesquisarTrabalhosGraduacao() {
	parent.mainFrame.location = 'index.asp?content=trabalho_graduacao' + getGlobalUrlParams();
}

function FiltroAnoValido(anoInicial, anoFinal, tipoFiltro) {
	if (anoInicial > 0) {
		var iAnoInicial;
		try {
			iAnoInicial = parseInt(anoInicial, 10);
			if (isNaN(iAnoInicial)) {
				throw "";
			}
		} catch (e) {
            alert(getTermo(global_frame.iIdioma, 9509, 'Ano inválido.', 0) + " " + getTermo(global_frame.iIdioma, 9519, 'Informe um ano entre 1900 e 2100.', 0));
			return false;
		}

		if (iAnoInicial < 1900 || iAnoInicial > 2100) {
            alert(getTermo(global_frame.iIdioma, 9509, 'Ano inválido.', 0) + " " + getTermo(global_frame.iIdioma, 9519, 'Informe um ano entre 1900 e 2100.', 0));
			return false;
		}
	}

	if (tipoFiltro == "3") {
		if (anoFinal > 0) {
			var iAnoFinal;
			try {
				iAnoFinal = parseInt(anoFinal, 10);
				if (isNaN(iAnoFinal)) {
					throw "";
				}
			} catch (e) {
                alert(getTermo(global_frame.iIdioma, 9509, 'Ano inválido.', 0) + " " + getTermo(global_frame.iIdioma, 9519, 'Informe um ano entre 1900 e 2100.', 0));
				return false;
			}

			if (iAnoFinal < 1900 || iAnoFinal > 2100) {
                alert(getTermo(global_frame.iIdioma, 9509, 'Ano inválido.', 0) + " " + getTermo(global_frame.iIdioma, 9519, 'Informe um ano entre 1900 e 2100.', 0));
				return false;
			} else {
				if (iAnoInicial > 0) {
                    if (iAnoFinal < iAnoInicial) {
                        alert(getTermo(global_frame.iIdioma, 5193, 'Ano final inválido.', 0) + " " + getTermo(global_frame.iIdioma, 9520, 'Informe um ano maior do que o ano inicial.', 0));
						return false;
					}
				} else {
                    alert(getTermo(global_frame.iIdioma, 5192, 'Ano inicial inválido.', 0));
					return false;
				}
			}
		}
	}
	return true;
}

function pesquisarTitulosTrabalhoGraduacao(curso, anoInicial, anoFinal, tipoFiltro) {
	global_frame.codigoCurso = curso;
	global_frame.anoInicial = anoInicial;
	global_frame.anoFinal = anoFinal;
	global_frame.tipoFiltro = tipoFiltro;
	var url = "index.asp?busca_customizada=ita&modo_busca=rapida&content=resultado&submeteu=rapida&rapida_campo=palavra_chave";
	parent.mainFrame.location = url + getGlobalUrlParams();
}