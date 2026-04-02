/*****************************************************************/
/*********** ATUALIZA OS COMBOS DE DADOS OPCIONAIS ***************/
/******* A REQUISIÇÃO E A ATUALIZAÇÃO É FEITA COM AJAX ***********/
/*****************************************************************/

var http_request_DadosOpc = false;
var atual = '';
var mobile = "0";

function atualizaDadosOpc(servidor, usuario, obra, cbAtual, origem_mobile, tipo_Obra) {
    mobile = origem_mobile;
    tipoObra = tipo_Obra;
    atual = '';
	
	//ANO
	if (document.frm_dados_comp.ano != null) {
		if (document.frm_dados_comp.ano.value == 'nenhum') {
			var ano = '';
		} else {
			var ano = document.frm_dados_comp.ano.value;
		}
	} else {
		var ano = '';
	}
	//VOLUME
	if (document.frm_dados_comp.volume != null) {
		if (document.frm_dados_comp.volume.value == 'nenhum') {
			var volume = '';
		} else {
			var volume = document.frm_dados_comp.volume.value;
		}
	} else {
		var volume = '';
	}
	//NUMERO
	if (document.frm_dados_comp.edicao != null) {
		if (document.frm_dados_comp.edicao.value == 'nenhum') {
			var numero = '';
		} else {
			var numero = RemoveAcentos(document.frm_dados_comp.edicao.value);
		}
	} else {
		var numero = '';
	}
	//SUPORTE
	if (document.frm_dados_comp.suporte != null) {
		if (document.frm_dados_comp.suporte.value == 'nenhum') {
			var suporte = '';
		} else {
			var suporte = document.frm_dados_comp.suporte.value;
		}
	} else {
		var suporte = '';
	}
	//BIBLIOTECA
	if (document.frm_dados_comp.biblioteca != null) {
		if (document.frm_dados_comp.biblioteca.value == 'nenhum') {
			var biblioteca = '';
		} else {
			var biblioteca = document.frm_dados_comp.biblioteca.value;
		}
	} else {
		var biblioteca = '';
	}
	//CHAMA A ATUALIZAÇÃO DOS COMBOS
	requestDadosOpc(servidor,usuario,obra,ano,volume,numero,suporte,biblioteca,cbAtual);
}

function requestDadosOpc(servidor,usuario,obra,ano,volume,numero,suporte,biblioteca,cbAtual) {
	var idioma = 0;
	if (parent.parent.hiddenFrame != null) {
		idioma = parent.parent.hiddenFrame.iIdioma;
	}

	http_request_DadosOpc = false;

	if (window.XMLHttpRequest) { // Mozilla, Safari,...
		http_request_DadosOpc = new XMLHttpRequest();
		if (http_request_DadosOpc.overrideMimeType) {
			http_request_DadosOpc.overrideMimeType('text/xml');
			// See note below about this line
		}
	} else if (window.ActiveXObject) { // IE
		try {
			http_request_DadosOpc = new ActiveXObject("Msxml2.XMLHTTP");
		} catch (e) {
			try {
				http_request_DadosOpc = new ActiveXObject("Microsoft.XMLHTTP");
			} catch (e) {}
		}
	}

	if (!http_request_DadosOpc) {
		alert(getTermo(idioma, 1270, 'Ocorreu um erro interno.', 0));
		return false;
	}
	
	atual = cbAtual;

    var url = "";
    if (mobile == "1") {
        url = '../asp/dados_opc_xml.asp?usuario=' + usuario + '&codigo=' + obra + '&ano=' + ano + '&volume=' + volume + '&numero=' + numero + '&suporte=' + suporte + '&biblioteca=' + biblioteca + '&servidor=' + servidor;
    } else {
        url = 'dados_opc_xml.asp?usuario=' + usuario + '&codigo=' + obra + '&ano=' + ano + '&volume=' + volume + '&numero=' + numero + '&suporte=' + suporte + '&biblioteca=' + biblioteca + '&servidor=' + servidor;
        parent.exibeLoadingPopup();
    }
	http_request_DadosOpc.onreadystatechange = processaDadosOpc;
	http_request_DadosOpc.open('GET', url, true);
	http_request_DadosOpc.send(null);
}

function processaDadosOpc() {
	var idioma = 0;
	if (parent.parent.hiddenFrame != null) {
		idioma = parent.parent.hiddenFrame.iIdioma;
	}
	
	var termo_nenhum = getTermo(idioma, 8012, 'Qualquer um', 0);
	if (parent.global_numero_serie != null) {
		if ((parent.global_numero_serie == 4134) || (parent.global_numero_serie == 5369)) {
			termo_nenhum = getTermo(idioma, 128, 'selecionar', 0); 
		}		
	}

	if (http_request_DadosOpc.readyState == 4) {
		if (http_request_DadosOpc.status == 200) {
	
			var xmldoc = http_request_DadosOpc.responseXML;
			
			//COMBO DE ANOS
			if ((document.frm_dados_comp.ano != null) && (document.frm_dados_comp.ano.options != null)) {
				if ((atual != 'ano') || ((document.frm_dados_comp.ano.value == '') || (document.frm_dados_comp.ano.value == 'nenhum'))) {

					if (parent.global_numero_serie != null) {
						if ((parent.global_numero_serie == 5859) || (parent.global_numero_serie == 5895) || (parent.global_numero_serie == 5896)) {
							if (tipoObra == 0) {
								termo_nenhum = getTermo(idioma, 128, 'selecionar', 0);
							} else {
								termo_nenhum = getTermo(idioma, 8012, 'Qualquer um', 0);
							}
						}
					}

					var exibeNenhum = false;
                    var AchouAno = false;
					var NenhumAno = false;
					var OldAno = document.frm_dados_comp.ano.value;
					var xAnos = xmldoc.getElementsByTagName('ANO');
					//Limpa o combo
					document.frm_dados_comp.ano.options.length = 0; 
					//Se possuir anos para serem exibidos inclui no combo
					if (xAnos.length > 0) {
						//Varre para cada ano
						for (var iNode = 0; iNode < xAnos.length; iNode++) {
							//Pega o ano atual e le atributos
							var xAno = xAnos[iNode];
							var DescAno = xAno.getAttribute('DESCRICAO');
							//Cria um novo item no combo (Texto, Valor)
							if (DescAno == '') {
								var OptAno = new Option(termo_nenhum,'nenhum');
								NenhumAno  = true;
							} else {
								var OptAno = new Option(DescAno,DescAno);
							}
							document.frm_dados_comp.ano.options.add(OptAno,document.frm_dados_comp.ano.options.length);
							if ((AchouAno == false) && (OldAno != '') && (OldAno == DescAno)) {
								AchouAno = true;
							}
						}
					}
					//Reposiciona o valor anterior no combo
					if (AchouAno == true) {
						document.frm_dados_comp.ano.value = OldAno;
					} else {
					    exibeNenhum = true;
                        document.frm_dados_comp.ano.value = 'nenhum';
					}
					//Desabilita combo quando há somente uma opção de seleção
					if (document.frm_dados_comp.ano.options.length > 1) {
					    if (mobile == "1") {
					        $("select[id='ano']").parent().parent().removeAttr('disabled');
					    }
					    document.frm_dados_comp.ano.disabled = false;
					    $('#ano + button.ui-multiselect').removeClass("ui-state-disabled");
					} else {					 
						if ((document.frm_dados_comp.ano.value == '') || (document.frm_dados_comp.ano.value == 'nenhum') || (exibeNenhum)) {
							if (NenhumAno == false) {
								var optionNenhumAno = new Option(termo_nenhum,'nenhum');
								document.frm_dados_comp.ano.options.add(optionNenhumAno,document.frm_dados_comp.ano.options.length);
								document.frm_dados_comp.ano.value = 'nenhum';
				            }
				            if (mobile == "1") {
				                $("select[id='ano']").parent().parent().attr('disabled', true);
				            }
				            $(document.frm_dados_comp.ano).attr('disabled', true);
                            $('#ano + button.ui-multiselect').addClass("ui-state-disabled");
						} else {
				            if (mobile == "1") {
				                $("select[id='ano']").parent().parent().removeAttr('disabled');
				            }
				            document.frm_dados_comp.ano.disabled = false;
				            $('#ano + button.ui-multiselect').removeClass("ui-state-disabled");
						}
					}
				}
			}
			
			//COMBO DE VOLUME
			if ((document.frm_dados_comp.volume != null) && (document.frm_dados_comp.volume.options != null)) {
			    if ((atual != 'vol') || ((document.frm_dados_comp.volume.value == '') || (document.frm_dados_comp.volume.value == 'nenhum'))) {

			    	if (parent.global_numero_serie != null) {
			    		if ((parent.global_numero_serie == 5859) || (parent.global_numero_serie == 5895) || (parent.global_numero_serie == 5896)) {
			    			termo_nenhum = getTermo(idioma, 128, 'selecionar', 0);
			    		}
			    	}

			    	var exibeNenhum = false;
					var AchouVol = false;
					var NenhumVol = false;
					var OldVol = document.frm_dados_comp.volume.value;
					var PrimValor = '';
					var xVolumes = xmldoc.getElementsByTagName('VOLUME');
					//Limpa o combo
					document.frm_dados_comp.volume.options.length = 0;
					//Se possuir volumes para serem exibidos inclui no combo
					if (xVolumes.length > 0) {
						//Varre para cada volume
						for (var iNode = 0; iNode < xVolumes.length; iNode++) {
							//Pega o volume atual e le atributos
							var xVolume = xVolumes[iNode];
							var DescVol = xVolume.getAttribute('DESCRICAO');
							//Pega o primeiro volume
							if (iNode == 0) {
								PrimValor = DescVol;
							}
							//Cria um novo item no combo (Texto, Valor)
							if (DescVol == '') {
								var OptVol = new Option(termo_nenhum,'nenhum');
								NenhumVol  = true;
							} else {
								var OptVol = new Option(DescVol,DescVol);
							}
							document.frm_dados_comp.volume.options.add(OptVol,document.frm_dados_comp.volume.options.length);
							if ((AchouVol == false) && (OldVol != '') && (OldVol == DescVol)) {
								AchouVol = true;
							}
						}
					}
					//Reposiciona o valor anterior no combo
                    if (AchouVol == true) {
						document.frm_dados_comp.volume.value = OldVol;
            		} else {
            		    exibeNenhum = true;
						document.frm_dados_comp.volume.value = 'nenhum';
					}
					//Desabilita combo quando há somente uma opção de seleção
					if (document.frm_dados_comp.volume.options.length > 1) {
					    if (mobile == "1") {
					        $("select[id='volume']").parent().parent().removeAttr('disabled');
					    }
					    document.frm_dados_comp.volume.disabled = false;
					    $('#volume + button.ui-multiselect').removeClass("ui-state-disabled");
		            } else {
		                if ((document.frm_dados_comp.volume.value == '') || (document.frm_dados_comp.volume.value == 'nenhum') || (exibeNenhum)) {
		                    if (NenhumVol == false) {
		                        var optionNenhumVol = new Option(termo_nenhum, 'nenhum');
		                        document.frm_dados_comp.volume.options.add(optionNenhumVol, document.frm_dados_comp.volume.options.length);
		                        document.frm_dados_comp.volume.value = 'nenhum';
		                    }
		                    if (mobile == "1") {
		                        $("select[id='volume']").parent().parent().attr('disabled', true);
		                    }
		                    document.frm_dados_comp.volume.disabled = true;
		                    $('#volume + button.ui-multiselect').addClass("ui-state-disabled");
						} else {
				            if (mobile == "1") {
				                $("select[id='volume']").parent().parent().removeAttr('disabled');
				            }
				            document.frm_dados_comp.volume.disabled = false;
				            $('#volume + button.ui-multiselect').removeClass("ui-state-disabled");
						}
					}
				}
			}
			
			//COMBO DE NUMERO
			if ((document.frm_dados_comp.edicao != null) && (document.frm_dados_comp.edicao.options != null)) {
				if ((atual != 'num') || ((document.frm_dados_comp.edicao.value == '') || (document.frm_dados_comp.edicao.value == 'nenhum'))) {

					if (parent.global_numero_serie != null) {
						if ((parent.global_numero_serie == 5859) || (parent.global_numero_serie == 5895) || (parent.global_numero_serie == 5896)) {
							if (tipoObra == 0) {
								termo_nenhum = getTermo(idioma, 128, 'selecionar', 0);
							} else {
								termo_nenhum = getTermo(idioma, 8012, 'Qualquer um', 0);
							}
						}
					}

					var exibeNenhum = false;
                    var AchouNum = false;
					var NenhumNum = false;
					var OldNum = document.frm_dados_comp.edicao.value;
					var PrimValor = '';
					var xNumeros = xmldoc.getElementsByTagName('NUMERO');
					//Limpa o combo
					document.frm_dados_comp.edicao.options.length = 0; 
					//Se possuir volumes para serem exibidos inclui no combo
					if (xNumeros.length > 0) {
						//Varre para cada ano
						for (var iNode = 0; iNode < xNumeros.length; iNode++) {
							//Pega o ano atual e le atributos
							var xNumero = xNumeros[iNode];
							var DescNum = xNumero.getAttribute('DESCRICAO');
							//Pega o primeiro número
							if (iNode == 0) {
								PrimValor = DescNum;
							}
							//Cria um novo item no combo (Texto, Valor)
							if (DescNum == '') {
								var OptNum = new Option(termo_nenhum,'nenhum');
								NenhumNum  = true;
							} else {
								var OptNum = new Option(DescNum,DescNum);
							}
							document.frm_dados_comp.edicao.options.add(OptNum,document.frm_dados_comp.edicao.options.length);
							if ((AchouNum == false) && (OldNum != '') && (OldNum == DescNum)) {
								AchouNum = true;
							}
						}
					}
					//Reposiciona o valor anterior no combo
					if (AchouNum == true) {
						document.frm_dados_comp.edicao.value = OldNum;
		            } else {
		                exibeNenhum = true;
						document.frm_dados_comp.edicao.value = 'nenhum';
					}
					//Desabilita combo quando há somente uma opção de seleção
					if (document.frm_dados_comp.edicao.options.length > 1) {
					    if (mobile == "1") {
					        $("select[id='edicao']").parent().parent().removeAttr('disabled');
					    }
					    document.frm_dados_comp.edicao.disabled = false;
					    $('#edicao + button.ui-multiselect').removeClass("ui-state-disabled");
					} else {
						if ((document.frm_dados_comp.edicao.value == '') || (document.frm_dados_comp.edicao.value == 'nenhum') || (exibeNenhum)) {
							if (NenhumNum == false) {
								var optionNenhumNum = new Option(termo_nenhum,'nenhum');
								document.frm_dados_comp.edicao.options.add(optionNenhumNum,document.frm_dados_comp.edicao.options.length);
								document.frm_dados_comp.edicao.value = 'nenhum';
				            }
				            if (mobile == "1") {
				                $("select[id='edicao']").parent().parent().attr('disabled', true);
				            }
				            document.frm_dados_comp.edicao.disabled = true;
				            $('#edicao + button.ui-multiselect').addClass("ui-state-disabled");
						} else {
			                if (mobile == "1") {
			                    $("select[id='edicao']").parent().parent().removeAttr('disabled');
			                }
			                document.frm_dados_comp.edicao.disabled = false;
			                $('#edicao + button.ui-multiselect').removeClass("ui-state-disabled");
						}
					}
				}
			}
			
			//COMBO DE SUPORTE
			if ((document.frm_dados_comp.suporte != null) && (document.frm_dados_comp.suporte.options != null)) {
				if ((atual != 'sup') || ((document.frm_dados_comp.suporte.value == '') || (document.frm_dados_comp.suporte.value == 'nenhum'))) {

					if ((parent.global_numero_serie == 5859) || (parent.global_numero_serie == 5895) || (parent.global_numero_serie == 5896)) {
						termo_nenhum = getTermo(idioma, 8012, 'Qualquer um', 0);
					}

					var exibeNenhum = false;
                    var AchouSup = false;
					var NenhumSup = false;
					var OldSup = document.frm_dados_comp.suporte.value;
					var xSuportes = xmldoc.getElementsByTagName('SUPORTE');
					//Limpa o combo
					document.frm_dados_comp.suporte.options.length = 0; 
					//Se possuir suportes para serem exibidos inclui no combo
					if (xSuportes.length > 0) {
						//Varre para cada suporte
						for (var iNode = 0; iNode < xSuportes.length; iNode++) {
							//Pega o suporte atual e le atributos
							var xSuporte = xSuportes[iNode];
							var CodSup = xSuporte.getAttribute('CODIGO');
							var DescSup = xSuporte.getAttribute('DESCRICAO');
							//Cria um novo item no combo (Texto, Valor)
							if (CodSup == '') {
								var OptSup = new Option(termo_nenhum,'nenhum');
								NenhumSup  = true;
							} else {
								var OptSup = new Option(DescSup,CodSup);
							}
							document.frm_dados_comp.suporte.options.add(OptSup,document.frm_dados_comp.suporte.options.length);
							if ((AchouSup == false) && (OldSup != '') && (OldSup == CodSup)) {
								AchouSup = true;
							}
						}
					}
					//Reposiciona o valor anterior no combo
					if (AchouSup == true) {
						document.frm_dados_comp.suporte.value = OldSup;
		            } else {
		                exibeNenhum = true;
						document.frm_dados_comp.suporte.value = 'nenhum';
					}
					//Desabilita combo quando há somente uma opção de seleção
					if (document.frm_dados_comp.suporte.options.length > 1) {
					    if (mobile == "1") {
					        $("select[id='suporte']").parent().parent().removeAttr('disabled');
					    }
					    document.frm_dados_comp.suporte.disabled = false;
					    $('#suporte + button.ui-multiselect').removeClass("ui-state-disabled");
					} else {
		                if ((document.frm_dados_comp.suporte.value == '') || (document.frm_dados_comp.suporte.value == 'nenhum') || (exibeNenhum)) {
							if (NenhumSup == false) {
								var optionNenhumSup = new Option(termo_nenhum,'nenhum');
								document.frm_dados_comp.suporte.options.add(optionNenhumSup,document.frm_dados_comp.suporte.options.length);
								document.frm_dados_comp.suporte.value = 'nenhum';
							}
				            if (mobile == "1") {
				                $("select[id='suporte']").parent().parent().attr('disabled', true);
				            }
				            document.frm_dados_comp.suporte.disabled = true;
				            $('#suporte + button.ui-multiselect').addClass("ui-state-disabled");
                        } else {
                            if (mobile == "1") {
                                $("select[id='suporte']").parent().parent().removeAttr('disabled');
                            }
                            document.frm_dados_comp.suporte.disabled = false;
                            $('#suporte + button.ui-multiselect').removeClass("ui-state-disabled");
						}
					}
				}
			}
			
			//COMBO DE BIBLIOTECAS
			if ((document.frm_dados_comp.biblioteca != null) && (document.frm_dados_comp.biblioteca.options != null)) {
				if ((atual != 'bib') || ((document.frm_dados_comp.biblioteca.value == '') || (document.frm_dados_comp.biblioteca.value == 'nenhum'))) {

					if ((parent.global_numero_serie == 5859) || (parent.global_numero_serie == 5895) || (parent.global_numero_serie == 5896)) {
						termo_nenhum = getTermo(idioma, 8012, 'Qualquer um', 0);
					}

					var AchouBib = false;
					var NenhumBib = false;
					var PrimValor = '';
					var OldBib = document.frm_dados_comp.biblioteca.value;
					var xBibliotecas = xmldoc.getElementsByTagName('BIBLIOTECA');
					//Limpa o combo
					document.frm_dados_comp.biblioteca.options.length = 0; 
					//Se possuir volumes para serem exibidos inclui no combo
					if (xBibliotecas.length > 0) {
						//Varre para cada ano
						for (var iNode = 0; iNode < xBibliotecas.length; iNode++) {
							//Pega o ano atual e le atributos
							var xBiblioteca = xBibliotecas[iNode];
							var CodBib = xBiblioteca.getAttribute('CODIGO');
							var DescBib = xBiblioteca.getAttribute('DESCRICAO');
							//Pega a primeira bibioteca
							if (iNode == 0) {
								PrimValor = CodBib;	
							}
							//Cria um novo item no combo (Texto, Valor)
							if (CodBib == '') {
								var OptBib = new Option(termo_nenhum,'nenhum');
								NenhumBib  = true;
							} else {
								var OptBib = new Option(DescBib,CodBib);
							}
							document.frm_dados_comp.biblioteca.options.add(OptBib,document.frm_dados_comp.biblioteca.options.length);
							if ((AchouBib == false) && (OldBib != '') && (OldBib == CodBib)) {
								AchouBib = true;
							}
						}
					}
					//Reposiciona o valor anterior no combo
					if (AchouBib == true) {
						document.frm_dados_comp.biblioteca.value = OldBib;
					} else {
						if (PrimValor == '') {
							document.frm_dados_comp.biblioteca.value = 'nenhum';
						} else {
							document.frm_dados_comp.biblioteca.value = PrimValor;
						}
					}
					//Desabilita combo quando há somente uma opção de seleção
					//if (document.frm_dados_comp.biblioteca.options.length > 1) {
					//    document.frm_dados_comp.biblioteca.disabled = false;
					//    $('#biblioteca + button.ui-multiselect').removeClass("ui-state-disabled");
					//} else {
		            //    document.frm_dados_comp.biblioteca.disabled = true;
		            //    $('#biblioteca + button.ui-multiselect').addClass("ui-state-disabled");
					//}
				}
            }
            
            //Refresh nos combos, para carregar as aletrações
			$('select').multiselect().multiselect("refresh");
  		
        } else {
			alert(getTermo(idioma, 1271, 'Ocorreu um erro durante a conexão com o servidor.', 0));			
		}

		if (mobile != "1") {
		    parent.fechaLoadingPopup();
		}
	}
}

function RemoveAcentos(strAccents) {
	var strAccents = strAccents.split('');
	var strAccentsOut = new Array();
	var strAccentsLen = strAccents.length;
	var accents = 'ÀÁÂÃÄÅàáâãäåÒÓÔÕÕÖØòóôõöøÈÉÊËèéêëðÇçÐÌÍÎÏìíîïÙÚÛÜùúûüÑñŠšŸÿýŽž';
	var accentsOut = "AAAAAAaaaaaaOOOOOOOooooooEEEEeeeeeCcDIIIIiiiiUUUUuuuuNnSsYyyZz";
	for (var y = 0; y < strAccentsLen; y++) {
		if (accents.indexOf(strAccents[y]) != -1) {
			strAccentsOut[y] = accentsOut.substr(accents.indexOf(strAccents[y]), 1);
		} else
			strAccentsOut[y] = strAccents[y];
	}
	strAccentsOut = strAccentsOut.join('');
	return strAccentsOut;
}