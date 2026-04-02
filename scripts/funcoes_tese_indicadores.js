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

function LinkIndicadores(modo) {
    parent.mainFrame.location = "index." + ext + "?modo_busca=" + modo + "&content=indicadores&veio_de=menu" + getGlobalUrlParams();
}

function LinkTeseDissertacao() {
    parent.mainFrame.location = "index." + ext + "?content=tese_dissertacao&veio_de=menu" + getGlobalUrlParams();
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
                        alert(getTermo(global_frame.iIdioma, 9509, 'Ano inválido.', 0) + " " + getTermo(global_frame.iIdioma, 9519, 'Informe um ano entre 1900 e 2100.', 0));
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

function pesquisarTitulosTesePorPrograma(material, programa, anoInicial, anoFinal, tipoFiltro) {
	pesquisarTitulosTesePorProgramaArea(material, programa, null, anoInicial, anoFinal, tipoFiltro);
}

function pesquisarTitulosTesePorProgramaArea(material, programa, area, anoInicial, anoFinal, tipoFiltro) {
	global_frame.codigoMaterial = material;
    global_frame.codigoPrograma = programa;
    if (area) {
        global_frame.codigoArea = area;
    } else {
        global_frame.codigoArea = 0;
    }
    global_frame.tipoFiltro = tipoFiltro;
    global_frame.anoInicial = anoInicial;
	global_frame.anoFinal = anoFinal;
    var url = "index.asp?busca_customizada=ita&modo_busca=rapida&content=resultado&submeteu=rapida&rapida_campo=palavra_chave";
    parent.mainFrame.location = url + getGlobalUrlParams();
}

function adicionarParametrosIta() {
	var parametrosIta = "";
	if (global_frame) {
		if (global_frame.tipoBuscaRegistro > 0) {
			parametrosIta = parametrosIta + "&tipoBuscaRegistro=" + global_frame.tipoBuscaRegistro;
		}
		if (global_frame.codigoCurso > 0) {
			parametrosIta = parametrosIta + "&curso=" + global_frame.codigoCurso;
		}
		if (global_frame.codigoMaterial > 0) {
			parametrosIta = parametrosIta + "&material=" + global_frame.codigoMaterial;
		}
		if (global_frame.codigoPrograma != 0) {
			parametrosIta = parametrosIta + "&programa=" + global_frame.codigoPrograma;
		}
		if (global_frame.codigoArea != 0) {
			parametrosIta = parametrosIta + "&area=" + global_frame.codigoArea;
		}
		if (global_frame.anoInicial > 0) {
			parametrosIta = parametrosIta + "&anoInicial=" + global_frame.anoInicial;
		}
		if (global_frame.anoFinal > 0) {
			parametrosIta = parametrosIta + "&anoFinal=" + global_frame.anoFinal;
		}
		if (global_frame.tipoFiltro > 0) {
			parametrosIta = parametrosIta + "&tipoFiltro=" + global_frame.tipoFiltro;
		}
	}
	return parametrosIta;
}