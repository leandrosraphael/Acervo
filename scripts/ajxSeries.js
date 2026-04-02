/*****************************************************************/
/********* ATUALIZA OS COMBOS DE SERIES EM SUGESTÕES *************/
/******* A REQUISIÇÃO E A ATUALIZAÇÃO É FEITA COM AJAX ***********/
/*****************************************************************/
var http_request_serie = false;
var select_serie = -1;
var form = null;

function atualizaSerie(servidor, serie_default, form) {
    this.form = form;
    //CURSO
    if (form.curso != null) {
        if (form.curso.value == '') {
            var curso = '-1';
        } else {
            var curso = form.curso.value;
        }
    } else {
        var curso = '-1';
    }
    //SERIE
    if (form.serie != null) {
        //Quando não há curso selecionado, desabilita o combo de séries
        if (curso == '-1') {
            form.serie.disabled = true;
            $('#serie + button.ui-multiselect').Attr("disabled", "disabled");
            $('#serie + button.ui-multiselect').addClass("ui-state-disabled");
        } else {
            form.serie.disabled = false;
            $('#serie + button.ui-multiselect').removeAttr("disabled");
            $('#serie + button.ui-multiselect').removeClass("ui-state-disabled");
        }
        //CHAMA A ATUALIZAÇÃO DOS COMBOS
        requestSerie(servidor, curso, serie_default);
    }
}

function requestSerie(servidor,curso,serie_default) {
	var idioma = 0;
	if (parent.hiddenFrame != null) {
		idioma = parent.hiddenFrame.iIdioma;
	}
	else
	if (parent.parent.hiddenFrame != null) {
		idioma = parent.parent.hiddenFrame.iIdioma;
	}

	http_request_serie = false;

	if (window.XMLHttpRequest) { // Mozilla, Safari,...
		http_request_serie = new XMLHttpRequest();
		if (http_request_serie.overrideMimeType) {
			http_request_serie.overrideMimeType('text/xml');
			// See note below about this line
		}
	} else if (window.ActiveXObject) { // IE
		try {
			http_request_serie = new ActiveXObject("Msxml2.XMLHTTP");
		} catch (e) {
			try {
				http_request_serie = new ActiveXObject("Microsoft.XMLHTTP");
			} catch (e) {}
		}
	}

	if (!http_request_serie) {
		alert(getTermo(idioma, 1270, 'Ocorreu um erro interno.', 0));
		return false;
	}
	
	select_serie = serie_default;
	
	var url = "";
	
	if (parent != null) {
		if (form == document.frm_sugestao) {
			url = 'serie_xml.asp?curso='+curso+'&servidor='+servidor;
			parent.exibeLoadingPopup();
		}
		else
		if (form == document.frm_bib_curso) {
			url = 'asp/serie_xml.asp?curso='+curso+'&servidor='+servidor;
			exibeLoadingPopup();
		}
	}

	http_request_serie.onreadystatechange = processaSerie;
	http_request_serie.open('GET', url, true);
	http_request_serie.send(null);
}

function processaSerie() {
	var idioma = 0;
	if (parent.hiddenFrame != null) {
		idioma = parent.hiddenFrame.iIdioma;
	}
	else
	if (parent.parent.hiddenFrame != null) {
		idioma = parent.parent.hiddenFrame.iIdioma;
	}

	if (http_request_serie.readyState == 4) {
		if (http_request_serie.status == 200) {
		
			var xmldoc = http_request_serie.responseXML;
			
			//COMBO DE SERIES
			if ((form.serie != null) && (form.serie.options != null)) {
				var AchouSerie = false;
				var OldSerie = form.serie.value;
				var xSeries = xmldoc.getElementsByTagName('REGISTRO');
				//Limpa o combo
				form.serie.options.length = 0; 
				//Cria um novo item no combo (Texto, Valor)
				if (form == document.frm_bib_curso) {
					var optionNenhumaSerie = new Option(getTermo(idioma, 3827, 'Todas', 0),'-1');
				}
				else{
					var optionNenhumaSerie = new Option('---------------------------','-1');
				}
				
				form.serie.options.add(optionNenhumaSerie,form.serie.options.length);
				//Se possuir series para serem exibidos inclui no combo
				if (xSeries.length > 0) {
					//Varre para cada ano
					for (var iNode = 0; iNode < xSeries.length; iNode++) {
						//Pega o ano atual e le atributos
		            	var xSerie = xSeries[iNode];
						var CodSerie = xSerie.getAttribute('CODIGO');
						var DescSerie = xSerie.getAttribute('DESCRICAO');
						//Cria um novo item no combo (Texto, Valor)
						var OptSerie = new Option(DescSerie,CodSerie);
						form.serie.options.add(OptSerie,form.serie.options.length);
						if ((AchouSerie == false) && (OldSerie != '') && (OldSerie != 'nenhum') && (OldSerie == CodSerie)) {
							AchouSerie = true;
						}
					}
				}
				//Reposiciona o valor anterior no combo
				if (select_serie != -1) {
					form.serie.value = select_serie;
				} else {
					if (AchouSerie == true) {
						form.serie.value = OldSerie;
					} else {
						form.serie.value = '-1';
					}
				}		
				
				if (parent.parent.hiddenFrame != null) {
					if (form == document.frm_sugestao) {
						if (parent.parent.hiddenFrame.AQ_CURSO == -1) {
							parent.parent.hiddenFrame.AQ_SERIE = -1;
						} else {
							parent.parent.hiddenFrame.AQ_SERIE = form.serie.value;
						}
					}
                }

				$(form.serie).multiselect().multiselect("refresh");
			}
			
			select_serie = -1;
			
			if (parent != null) {
				if (form == document.frm_sugestao) {
					parent.fechaLoadingPopup();
				}
				else
				if (form == document.frm_bib_curso) {
					fechaLoadingPopup();
				}
			}
		} else {
			select_serie = -1;
			if (form == document.frm_sugestao) {
				parent.fechaLoadingPopup();
			}
			else
			if (form == document.frm_bib_curso) {
				fechaLoadingPopup();
			}
			
			alert(getTermo(idioma, 1271, 'Ocorreu um erro durante a conexão com o servidor.', 0));
		}
	}
	this.form = null;
}