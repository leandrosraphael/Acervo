//**************************
// ROTINAS AJAX
//**************************

function posTop() {
	var win = self;
	var doc = document;
	return typeof win.pageYOffset != 'undefined' ? win.pageYOffset:doc.documentElement && doc.documentElement.scrollTop? doc.documentElement.scrollTop: doc.body.scrollTop?doc.body.scrollTop:0;
}

function ajxStartLoad(URL) {
	div = document.getElementById('dvLoad');
	if (div != null) {
		div.style.top = posTop() + 150;
		div.style.display = 'block';
	}
	if (URL != "") {
		frm = document.getElementById('ajxFrame');
		frm.src = URL;
	}
}

function ajxFinishLoad(HTML) {
	div = document.getElementById('ajxDiv');
	if (div != null) {
		div.innerHTML = HTML;
	}
	div = document.getElementById('dvLoad');
	if (div != null) {
		div.style.display = 'none';
	}
	window.scroll(0,0);
}

function setFocus(obj) {
	//SETA FOCO NO EDIT QUANDO EXISTE NA TELA
	var edBusca = document.getElementById(obj);
	if (edBusca != null) {
		edBusca.focus();
	}
}

//**************************
// ROTINAS VALIDAÇÃO BUSCA
//**************************

function ValidaForm() {
	if (document.frm_pesquisa.dados_outro != null) {
		var tam_dados = document.frm_pesquisa.dados_outro.value.length;
		if (tam_dados == 0) {
			alert("O campo refinar não pode ser vazio!");
			document.frm_pesquisa.dados_outro.focus();
			return false;
		} else {
			document.frm_pesquisa.submit.enabled = false;
			ajxStartLoad("");
			return true;
		}	
	} else { 
		var tam_dados = document.frm_pesquisa.dados.value.length;
		//VALIDA CAMPO OPCIONAL 1
		if (document.frm_pesquisa.tipo_campo1 != null) {
			if (document.frm_pesquisa.tipo_campo1.value == 'ALFA') {
				if (document.frm_pesquisa.campo1 != null) {
					tam_dados += document.frm_pesquisa.campo1.value.length;
				}
			} else if (document.frm_pesquisa.tipo_campo1.value == 'LOGICO') {
				if (document.frm_pesquisa.campo1 != null) {
					if (document.frm_pesquisa.campo1.value != '') {
						tam_dados += 1;
					}
				}
			} else if (document.frm_pesquisa.tipo_campo1.value == 'NUM') {
				if (document.frm_pesquisa.cb_campo1.value == '3') {
					if (document.frm_pesquisa.campo1_ini.value == '') {
						alert('Preencha o intervalo inicial!');
						document.frm_pesquisa.campo1_ini.focus();
						return false;
					} else if (document.frm_pesquisa.campo1_fim.value == '') {
						alert('Preencha o intervalo final!');
						document.frm_pesquisa.campo1_fim.focus();
						return false;
					} else {
						tam_dados += 1;
					}
				} else {
					if (document.frm_pesquisa.campo1_ini.value != '') {
						tam_dados += 1;
					}
				}
			} 
			else if (document.frm_pesquisa.tipo_campo1.value == 'TABELA') 
			{
				tam_dados += document.frm_pesquisa.cb_campo1.value;
			} 
			else if (document.frm_pesquisa.tipo_campo1.value == 'DATA') {
				if (document.frm_pesquisa.campo1_ini_d.value.length == 1) {
					document.frm_pesquisa.campo1_ini_d.value = '0' + document.frm_pesquisa.campo1_ini_d.value;
				}
				if (document.frm_pesquisa.campo1_ini_m.value.length == 1) {
					document.frm_pesquisa.campo1_ini_m.value = '0' + document.frm_pesquisa.campo1_ini_m.value;
				}
				if (document.frm_pesquisa.campo1_ini_a.value.length == 1) {
					document.frm_pesquisa.campo1_ini_a.value = '200' + document.frm_pesquisa.campo1_ini_a.value;
				} else if (document.frm_pesquisa.campo1_ini_a.value.length == 2) {
					if (document.frm_pesquisa.campo1_ini_a.value > 50) {
						document.frm_pesquisa.campo1_ini_a.value = '20' + document.frm_pesquisa.campo1_ini_a.value;
					} else {
						document.frm_pesquisa.campo1_ini_a.value = '19' + document.frm_pesquisa.campo1_ini_a.value;
					}
				} else if (document.frm_pesquisa.campo1_ini_a.value.length == 3) {
					document.frm_pesquisa.campo1_ini_a.value = '0' + document.frm_pesquisa.campo1_ini_a.value;
				}

				if (document.frm_pesquisa.cb_campo1.value == '3') {
					if (document.frm_pesquisa.campo1_fim_d.value.length == 1) {
						document.frm_pesquisa.campo1_fim_d.value = '0' + document.frm_pesquisa.campo1_fim_d.value;
					}
					if (document.frm_pesquisa.campo1_fim_m.value.length == 1) {
						document.frm_pesquisa.campo1_fim_m.value = '0' + document.frm_pesquisa.campo1_fim_m.value;
					}
					if (document.frm_pesquisa.campo1_fim_a.value.length == 1) {
						document.frm_pesquisa.campo1_fim_a.value = '200' + document.frm_pesquisa.campo1_fim_a.value;
					} else if (document.frm_pesquisa.campo1_fim_a.value.length == 2) {
						if (document.frm_pesquisa.campo1_fim_a.value > 50) {
							document.frm_pesquisa.campo1_fim_a.value = '20' + document.frm_pesquisa.campo1_fim_a.value;
						} else {
							document.frm_pesquisa.campo1_fim_a.value = '19' + document.frm_pesquisa.campo1_fim_a.value;
						}
					} else if (document.frm_pesquisa.campo1_fim_a.value.length == 3) {
						document.frm_pesquisa.campo1_fim_a.value = '0' + document.frm_pesquisa.campo1_fim_a.value;
					}

					if ((document.frm_pesquisa.campo1_ini_d.value == '')||(document.frm_pesquisa.campo1_ini_d.value > 31)) {
						alert('Digite um dia válido!');
						document.frm_pesquisa.campo1_ini_d.focus();
						return false;
					} else if ((document.frm_pesquisa.campo1_ini_m.value == '')||(document.frm_pesquisa.campo1_ini_m.value > 12)) {
						alert('Digite um mês válido!');
						document.frm_pesquisa.campo1_ini_m.focus();
						return false;
					} else if (document.frm_pesquisa.campo1_ini_a.value == '') {
						alert('Digite um ano válido!');
						document.frm_pesquisa.campo1_ini_a.focus();
						return false;
					} else if ((document.frm_pesquisa.campo1_fim_d.value == '')||(document.frm_pesquisa.campo1_fim_d.value > 31)) {
						alert('Digite um dia válido!');
						document.frm_pesquisa.campo1_fim_d.focus();
						return false;
					} else if ((document.frm_pesquisa.campo1_fim_m.value == '')||(document.frm_pesquisa.campo1_fim_m.value > 12)) {
						alert('Digite um mês válido!');
						document.frm_pesquisa.campo1_fim_m.focus();
						return false;
					} else if (document.frm_pesquisa.campo1_fim_a.value == '') {
						alert('Digite um ano válido!');
						document.frm_pesquisa.campo1_fim_a.focus();
						return false;
					} else {
						tam_dados += 1;
					}
				} else {
					if ((document.frm_pesquisa.campo1_ini_d.value != '') || (document.frm_pesquisa.campo1_ini_m.value != '') || 
 					    (document.frm_pesquisa.campo1_ini_a.value != '')) {
						if ((document.frm_pesquisa.campo1_ini_d.value > 31) || (document.frm_pesquisa.campo1_ini_d.value == '')) {
							alert('Digite um dia válido!');
							document.frm_pesquisa.campo1_ini_d.focus();
							return false;
						} else if ((document.frm_pesquisa.campo1_ini_m.value > 12) || (document.frm_pesquisa.campo1_ini_m.value == '')) {
							alert('Digite um mês válido!');
							document.frm_pesquisa.campo1_ini_m.focus();
							return false;
						} else if (document.frm_pesquisa.campo1_ini_a.value == '') {
							alert('Digite um ano válido!');
							document.frm_pesquisa.campo1_ini_a.focus();
							return false;
						} else {
							tam_dados += 1;
						}
					}
				}
			}
		}
		//VALIDA CAMPO OPCIONAL 2
		if (document.frm_pesquisa.tipo_campo2 != null) {
			if (document.frm_pesquisa.tipo_campo2.value == 'ALFA') {
				if (document.frm_pesquisa.campo2 != null) {
					tam_dados += document.frm_pesquisa.campo2.value.length;
				}
			} else if (document.frm_pesquisa.tipo_campo2.value == 'LOGICO') {
				if (document.frm_pesquisa.campo2 != null) {
					if (document.frm_pesquisa.campo2.value != '') {
						tam_dados += 1;
					}
				}
			} else if (document.frm_pesquisa.tipo_campo2.value == 'NUM') {
				if (document.frm_pesquisa.cb_campo2.value == '3') {
					if (document.frm_pesquisa.campo2_ini.value == '') {
						alert('Preencha o intervalo inicial!');
						document.frm_pesquisa.campo2_ini.focus();
						return false;
					} else if (document.frm_pesquisa.campo2_fim.value == '') {
						alert('Preencha o intervalo final!');
						document.frm_pesquisa.campo2_fim.focus();
						return false;
					} else {
						tam_dados += 1;
					}
				} else {
					if (document.frm_pesquisa.campo2_ini.value != '') {
						tam_dados += 1;
					}
				}				
			} 
			else if (document.frm_pesquisa.tipo_campo2.value == 'TABELA') 
			{
				tam_dados += document.frm_pesquisa.cb_campo2.value;
			} 			
			else if (document.frm_pesquisa.tipo_campo2.value == 'DATA') {
				if (document.frm_pesquisa.campo2_ini_d.value.length == 1) {
					document.frm_pesquisa.campo2_ini_d.value = '0' + document.frm_pesquisa.campo2_ini_d.value;
				}
				if (document.frm_pesquisa.campo2_ini_m.value.length == 1) {
					document.frm_pesquisa.campo2_ini_m.value = '0' + document.frm_pesquisa.campo2_ini_m.value;
				}
				if (document.frm_pesquisa.campo2_ini_a.value.length == 1) {
					document.frm_pesquisa.campo2_ini_a.value = '200' + document.frm_pesquisa.campo2_ini_a.value;
				} else if (document.frm_pesquisa.campo2_ini_a.value.length == 2) {
					if (document.frm_pesquisa.campo2_ini_a.value > 50) {
						document.frm_pesquisa.campo2_ini_a.value = '20' + document.frm_pesquisa.campo2_ini_a.value;
					} else {
						document.frm_pesquisa.campo2_ini_a.value = '19' + document.frm_pesquisa.campo2_ini_a.value;
					}
				} else if (document.frm_pesquisa.campo2_ini_a.value.length == 3) {
					document.frm_pesquisa.campo2_ini_a.value = '0' + document.frm_pesquisa.campo2_ini_a.value;
				}

				if (document.frm_pesquisa.cb_campo2.value == '3') {
					if (document.frm_pesquisa.campo2_fim_d.value.length == 1) {
						document.frm_pesquisa.campo2_fim_d.value = '0' + document.frm_pesquisa.campo2_fim_d.value;
					}
					if (document.frm_pesquisa.campo2_fim_m.value.length == 1) {
						document.frm_pesquisa.campo2_fim_m.value = '0' + document.frm_pesquisa.campo2_fim_m.value;
					}
					if (document.frm_pesquisa.campo2_fim_a.value.length == 1) {
						document.frm_pesquisa.campo2_fim_a.value = '200' + document.frm_pesquisa.campo2_fim_a.value;
					} else if (document.frm_pesquisa.campo2_fim_a.value.length == 2) {
						if (document.frm_pesquisa.campo2_fim_a.value > 50) {
							document.frm_pesquisa.campo2_fim_a.value = '20' + document.frm_pesquisa.campo2_fim_a.value;
						} else {
							document.frm_pesquisa.campo2_fim_a.value = '19' + document.frm_pesquisa.campo2_fim_a.value;
						}
					} else if (document.frm_pesquisa.campo2_fim_a.value.length == 3) {
						document.frm_pesquisa.campo2_fim_a.value = '0' + document.frm_pesquisa.campo2_fim_a.value;
					}
					
					if ((document.frm_pesquisa.campo2_ini_d.value == '')||(document.frm_pesquisa.campo2_ini_d.value > 31)) {
						alert('Digite um dia válido!');
						document.frm_pesquisa.campo2_ini_d.focus();
						return false;
					} else if ((document.frm_pesquisa.campo2_ini_m.value == '')||(document.frm_pesquisa.campo2_ini_m.value > 12)) {
						alert('Digite um mês válido!');
						document.frm_pesquisa.campo2_ini_m.focus();
						return false;
					} else if (document.frm_pesquisa.campo2_ini_a.value == '') {
						alert('Digite um ano válido!');
						document.frm_pesquisa.campo2_ini_a.focus();
						return false;
					} else if ((document.frm_pesquisa.campo2_fim_d.value == '')||(document.frm_pesquisa.campo2_fim_d.value > 31)) {
						alert('Digite um dia válido!');
						document.frm_pesquisa.campo2_fim_d.focus();
						return false;
					} else if ((document.frm_pesquisa.campo2_fim_m.value == '')||(document.frm_pesquisa.campo2_fim_m.value > 12)) {
						alert('Digite um mês válido!');
						document.frm_pesquisa.campo2_fim_m.focus();
						return false;
					} else if (document.frm_pesquisa.campo2_fim_a.value == '') {
						alert('Digite um ano válido!');
						document.frm_pesquisa.campo2_fim_a.focus();
						return false;
					} else {
						tam_dados += 1;
					}
				} else {
					if ((document.frm_pesquisa.campo2_ini_d.value != '') || (document.frm_pesquisa.campo2_ini_m.value != '') || 
 					    (document.frm_pesquisa.campo2_ini_a.value != '')) {
						if ((document.frm_pesquisa.campo2_ini_d.value > 31) || (document.frm_pesquisa.campo2_ini_d.value == '')) {
							alert('Digite um dia válido!');
							document.frm_pesquisa.campo2_ini_d.focus();
							return false;
						} else if ((document.frm_pesquisa.campo2_ini_m.value > 12) || (document.frm_pesquisa.campo2_ini_m.value == '')) {
							alert('Digite um mês válido!');
							document.frm_pesquisa.campo2_ini_m.focus();
							return false;
						} else if (document.frm_pesquisa.campo2_ini_a.value == '') {
							alert('Digite um ano válido!');
							document.frm_pesquisa.campo2_ini_a.focus();
							return false;
						} else {
							tam_dados += 1;
						}
					}
				}
			}
		}
		//VALIDA CAMPO OPCIONAL 3
		if (document.frm_pesquisa.tipo_campo3 != null) {
			if (document.frm_pesquisa.tipo_campo3.value == 'ALFA') {
				if (document.frm_pesquisa.campo3 != null) {
					tam_dados += document.frm_pesquisa.campo3.value.length;
				}
			} else if (document.frm_pesquisa.tipo_campo3.value == 'LOGICO') {
				if (document.frm_pesquisa.campo3 != null) {
					if (document.frm_pesquisa.campo3.value != '') {
						tam_dados += 1;
					}
				}
			} else if (document.frm_pesquisa.tipo_campo3.value == 'NUM') {
				if (document.frm_pesquisa.cb_campo3.value == '3') {
					if (document.frm_pesquisa.campo3_ini.value == '') {
						alert('Preencha o intervalo inicial!');
						document.frm_pesquisa.campo3_ini.focus();
						return false;
					} else if (document.frm_pesquisa.campo3_fim.value == '') {
						alert('Preencha o intervalo final!');
						document.frm_pesquisa.campo3_fim.focus();
						return false;
					} else {
						tam_dados += 1;
					}
				} else {
					if (document.frm_pesquisa.campo3_ini.value != '') {
						tam_dados += 1;
					}
				}
			} 
			else if (document.frm_pesquisa.tipo_campo3.value == 'TABELA') 
			{
				tam_dados += document.frm_pesquisa.cb_campo3.value;
			} 			
			else if (document.frm_pesquisa.tipo_campo3.value == 'DATA') {
				if (document.frm_pesquisa.campo3_ini_d.value.length == 1) {
					document.frm_pesquisa.campo3_ini_d.value = '0' + document.frm_pesquisa.campo3_ini_d.value;
				}
				if (document.frm_pesquisa.campo3_ini_m.value.length == 1) {
					document.frm_pesquisa.campo3_ini_m.value = '0' + document.frm_pesquisa.campo3_ini_m.value;
				}
				if (document.frm_pesquisa.campo3_ini_a.value.length == 1) {
					document.frm_pesquisa.campo3_ini_a.value = '200' + document.frm_pesquisa.campo3_ini_a.value;
				} else if (document.frm_pesquisa.campo3_ini_a.value.length == 2) {
					if (document.frm_pesquisa.campo3_ini_a.value > 50) {
						document.frm_pesquisa.campo3_ini_a.value = '20' + document.frm_pesquisa.campo3_ini_a.value;
					} else {
						document.frm_pesquisa.campo3_ini_a.value = '19' + document.frm_pesquisa.campo3_ini_a.value;
					}
				} else if (document.frm_pesquisa.campo3_ini_a.value.length == 3) {
					document.frm_pesquisa.campo3_ini_a.value = '0' + document.frm_pesquisa.campo3_ini_a.value;
				}
				
				if (document.frm_pesquisa.cb_campo3.value == '3') {
					if (document.frm_pesquisa.campo3_fim_d.value.length == 1) {
						document.frm_pesquisa.campo3_fim_d.value = '0' + document.frm_pesquisa.campo3_fim_d.value;
					}
					if (document.frm_pesquisa.campo3_fim_m.value.length == 1) {
						document.frm_pesquisa.campo3_fim_m.value = '0' + document.frm_pesquisa.campo3_fim_m.value;
					}
					if (document.frm_pesquisa.campo3_fim_a.value.length == 1) {
						document.frm_pesquisa.campo3_fim_a.value = '200' + document.frm_pesquisa.campo3_fim_a.value;
					} else if (document.frm_pesquisa.campo3_fim_a.value.length == 2) {
						if (document.frm_pesquisa.campo3_fim_a.value > 50) {
							document.frm_pesquisa.campo3_fim_a.value = '20' + document.frm_pesquisa.campo3_fim_a.value;
						} else {
							document.frm_pesquisa.campo3_fim_a.value = '19' + document.frm_pesquisa.campo3_fim_a.value;
						}
					} else if (document.frm_pesquisa.campo3_fim_a.value.length == 3) {
						document.frm_pesquisa.campo3_fim_a.value = '0' + document.frm_pesquisa.campo3_fim_a.value;
					}
					
					if ((document.frm_pesquisa.campo3_ini_d.value == '')||(document.frm_pesquisa.campo3_ini_d.value > 31)) {
						alert('Digite um dia válido!');
						document.frm_pesquisa.campo3_ini_d.focus();
						return false;
					} else if ((document.frm_pesquisa.campo3_ini_m.value == '')||(document.frm_pesquisa.campo3_ini_m.value > 12)) {
						alert('Digite um mês válido!');
						document.frm_pesquisa.campo3_ini_m.focus();
						return false;
					} else if (document.frm_pesquisa.campo3_ini_a.value == '') {
						alert('Digite um ano válido!');
						document.frm_pesquisa.campo3_ini_a.focus();
						return false;
					} else if ((document.frm_pesquisa.campo3_fim_d.value == '')||(document.frm_pesquisa.campo3_fim_d.value > 31)) {
						alert('Digite um dia válido!');
						document.frm_pesquisa.campo3_fim_d.focus();
						return false;
					} else if ((document.frm_pesquisa.campo3_fim_m.value == '')||(document.frm_pesquisa.campo3_fim_m.value > 12)) {
						alert('Digite um mês válido!');
						document.frm_pesquisa.campo3_fim_m.focus();
						return false;
					} else if (document.frm_pesquisa.campo3_fim_a.value == '') {
						alert('Digite um ano válido!');
						document.frm_pesquisa.campo3_fim_a.focus();
						return false;
					} else {
						tam_dados += 1;
					}
				} else {
					if ((document.frm_pesquisa.campo3_ini_d.value != '') || (document.frm_pesquisa.campo3_ini_m.value != '') || 
 					    (document.frm_pesquisa.campo3_ini_a.value != '')) {
						if ((document.frm_pesquisa.campo3_ini_d.value > 31) || (document.frm_pesquisa.campo3_ini_d.value == '')) {
							alert('Digite um dia válido!');
							document.frm_pesquisa.campo3_ini_d.focus();
							return false;
						} else if ((document.frm_pesquisa.campo3_ini_m.value > 12) || (document.frm_pesquisa.campo3_ini_m.value == '')) {
							alert('Digite um mês válido!');
							document.frm_pesquisa.campo3_ini_m.focus();
							return false;
						} else if (document.frm_pesquisa.campo3_ini_a.value == '') {
							alert('Digite um ano válido!');
							document.frm_pesquisa.campo3_ini_a.focus();
							return false;
						} else {
							tam_dados += 1;
						}
					}
				}
			}
		}
		//VALIDA CAMPO OPCIONAL 4
		if (document.frm_pesquisa.tipo_campo4 != null) {
			if (document.frm_pesquisa.tipo_campo4.value == 'ALFA') {
				if (document.frm_pesquisa.campo4 != null) {
					tam_dados += document.frm_pesquisa.campo4.value.length;
				}
			} else if (document.frm_pesquisa.tipo_campo4.value == 'LOGICO') {
				if (document.frm_pesquisa.campo4 != null) {
					if (document.frm_pesquisa.campo4.value != '') {
						tam_dados += 1;
					}
				}
			} else if (document.frm_pesquisa.tipo_campo4.value == 'NUM') {
				if (document.frm_pesquisa.cb_campo4.value == '3') {
					if (document.frm_pesquisa.campo4_ini.value == '') {
						alert('Preencha o intervalo inicial!');
						document.frm_pesquisa.campo4_ini.focus();
						return false;
					} else if (document.frm_pesquisa.campo4_fim.value == '') {
						alert('Preencha o intervalo final!');
						document.frm_pesquisa.campo4_fim.focus();
						return false;
					} else {
						tam_dados += 1;
					}
				} else {
					if (document.frm_pesquisa.campo4_ini.value != '') {
						tam_dados += 1;
					}
				}
			} 
			else if (document.frm_pesquisa.tipo_campo4.value == 'TABELA') 
			{
				tam_dados += document.frm_pesquisa.cb_campo4.value;
			} 			
			else if (document.frm_pesquisa.tipo_campo4.value == 'DATA') {
				if (document.frm_pesquisa.campo4_ini_d.value.length == 1) {
					document.frm_pesquisa.campo4_ini_d.value = '0' + document.frm_pesquisa.campo4_ini_d.value;
				}
				if (document.frm_pesquisa.campo4_ini_m.value.length == 1) {
					document.frm_pesquisa.campo4_ini_m.value = '0' + document.frm_pesquisa.campo4_ini_m.value;
				}
				if (document.frm_pesquisa.campo4_ini_a.value.length == 1) {
					document.frm_pesquisa.campo4_ini_a.value = '200' + document.frm_pesquisa.campo4_ini_a.value;
				} else if (document.frm_pesquisa.campo4_ini_a.value.length == 2) {
					if (document.frm_pesquisa.campo4_ini_a.value > 50) {
						document.frm_pesquisa.campo4_ini_a.value = '20' + document.frm_pesquisa.campo4_ini_a.value;
					} else {
						document.frm_pesquisa.campo4_ini_a.value = '19' + document.frm_pesquisa.campo4_ini_a.value;
					}
				} else if (document.frm_pesquisa.campo4_ini_a.value.length == 3) {
					document.frm_pesquisa.campo4_ini_a.value = '0' + document.frm_pesquisa.campo4_ini_a.value;
				}
				
				if (document.frm_pesquisa.cb_campo4.value == '3') {
					if (document.frm_pesquisa.campo4_fim_d.value.length == 1) {
						document.frm_pesquisa.campo4_fim_d.value = '0' + document.frm_pesquisa.campo4_fim_d.value;
					}
					if (document.frm_pesquisa.campo4_fim_m.value.length == 1) {
						document.frm_pesquisa.campo4_fim_m.value = '0' + document.frm_pesquisa.campo4_fim_m.value;
					}
					if (document.frm_pesquisa.campo4_fim_a.value.length == 1) {
						document.frm_pesquisa.campo4_fim_a.value = '200' + document.frm_pesquisa.campo4_fim_a.value;
					} else if (document.frm_pesquisa.campo4_fim_a.value.length == 2) {
						if (document.frm_pesquisa.campo4_fim_a.value > 50) {
							document.frm_pesquisa.campo4_fim_a.value = '20' + document.frm_pesquisa.campo4_fim_a.value;
						} else {
							document.frm_pesquisa.campo4_fim_a.value = '19' + document.frm_pesquisa.campo4_fim_a.value;
						}
					} else if (document.frm_pesquisa.campo4_fim_a.value.length == 3) {
						document.frm_pesquisa.campo4_fim_a.value = '0' + document.frm_pesquisa.campo4_fim_a.value;
					}
					
					if ((document.frm_pesquisa.campo4_ini_d.value == '')||(document.frm_pesquisa.campo4_ini_d.value > 31)) {
						alert('Digite um dia válido!');
						document.frm_pesquisa.campo4_ini_d.focus();
						return false;
					} else if ((document.frm_pesquisa.campo4_ini_m.value == '')||(document.frm_pesquisa.campo4_ini_m.value > 12)) {
						alert('Digite um mês válido!');
						document.frm_pesquisa.campo4_ini_m.focus();
						return false;
					} else if (document.frm_pesquisa.campo4_ini_a.value == '') {
						alert('Digite um ano válido!');
						document.frm_pesquisa.campo4_ini_a.focus();
						return false;
					} else if ((document.frm_pesquisa.campo4_fim_d.value == '')||(document.frm_pesquisa.campo4_fim_d.value > 31)) {
						alert('Digite um dia válido!');
						document.frm_pesquisa.campo4_fim_d.focus();
						return false;
					} else if ((document.frm_pesquisa.campo4_fim_m.value == '')||(document.frm_pesquisa.campo4_fim_m.value > 12)) {
						alert('Digite um mês válido!');
						document.frm_pesquisa.campo4_fim_m.focus();
						return false;
					} else if (document.frm_pesquisa.campo4_fim_a.value == '') {
						alert('Digite um ano válido!');
						document.frm_pesquisa.campo4_fim_a.focus();
						return false;
					} else {					
						tam_dados += 1;
					}
				} else {
					if ((document.frm_pesquisa.campo4_ini_d.value != '') || (document.frm_pesquisa.campo4_ini_m.value != '') || 
 					    (document.frm_pesquisa.campo4_ini_a.value != '')) {
						if ((document.frm_pesquisa.campo4_ini_d.value > 31) || (document.frm_pesquisa.campo4_ini_d.value == '')) {
							alert('Digite um dia válido!');
							document.frm_pesquisa.campo4_ini_d.focus();
							return false;
						} else if ((document.frm_pesquisa.campo4_ini_m.value > 12) || (document.frm_pesquisa.campo4_ini_m.value == '')) {
							alert('Digite um mês válido!');
							document.frm_pesquisa.campo4_ini_m.focus();
							return false;
						} else if (document.frm_pesquisa.campo4_ini_a.value == '') {
							alert('Digite um ano válido!');
							document.frm_pesquisa.campo4_ini_a.focus();
							return false;
						} else {
							tam_dados += 1;
						}
					}
				}
			}
		}
		if (document.frm_pesquisa.contexto != null) {
			var iCont = document.frm_pesquisa.contexto.value;
		} else {
			var iCont = 0;
		}
		if ((tam_dados == 0) && (iCont == 0)) {
			alert("Por favor, preencha pelo menos um campo de pesquisa!");
			document.frm_pesquisa.dados.focus();
			return false;
		} else {
			document.frm_pesquisa.submit.enabled = false;
			ajxStartLoad("");
			return true;
		}
	}		
}

function resetForm() {
	var edDados    = document.getElementById('dados');
	var cbObjeto   = document.getElementById('objeto');
	var cbContexto = document.getElementById('contexto');
	var ckImagem = document.getElementById('imagem');
	if (edDados != null) {
		edDados.value = '';
	}
	if (cbObjeto != null) {
		cbObjeto.value = 0;
	}
	if (cbContexto != null) {
		cbContexto.value = 0;
	}
	if (ckImagem != null) {
		ckImagem.checked = false;
	}
	if (document.frm_pesquisa.tipo_campo1 != null) {
		if (document.frm_pesquisa.tipo_campo1.value == 'ALFA') {
			if (document.frm_pesquisa.campo1 != null) {
				document.frm_pesquisa.campo1.value = '';
			}
		} else if (document.frm_pesquisa.tipo_campo1.value == 'LOGICO') {
			if (document.frm_pesquisa.campo1 != null) {
				document.frm_pesquisa.campo1.value = '';
			}
		} else if (document.frm_pesquisa.tipo_campo1.value == 'NUM') {
			if (document.frm_pesquisa.cb_campo1 != null) {
				document.frm_pesquisa.cb_campo1.value = 0;
			}
			if (document.frm_pesquisa.campo1_ini != null) {
				document.frm_pesquisa.campo1_ini.value = '';
			}
			if (document.frm_pesquisa.campo1_fim != null) {
				document.frm_pesquisa.campo1_fim.value = '';
			}
			var obj_span = MM_findObj('span_campo1');
			if (obj_span != null) {
				obj_span.style.display = 'none';
			}
		} else if (document.frm_pesquisa.tipo_campo1.value == 'DATA') {
			if (document.frm_pesquisa.cb_campo1 != null) {
				document.frm_pesquisa.cb_campo1.value = 0;
			}
			if (document.frm_pesquisa.campo1_ini_d != null) {
				document.frm_pesquisa.campo1_ini_d.value = '';
			}
			if (document.frm_pesquisa.campo1_ini_m != null) {
				document.frm_pesquisa.campo1_ini_m.value = '';
			}
			if (document.frm_pesquisa.campo1_ini_a != null) {
				document.frm_pesquisa.campo1_ini_a.value = '';
			}
			if (document.frm_pesquisa.campo1_fim_d != null) {
				document.frm_pesquisa.campo1_fim_d.value = '';
			}
			if (document.frm_pesquisa.campo1_fim_m != null) {
				document.frm_pesquisa.campo1_fim_m.value = '';
			}
			if (document.frm_pesquisa.campo1_fim_a != null) {
				document.frm_pesquisa.campo1_fim_a.value = '';
			}
			var obj_span = MM_findObj('span_campo1');
			if (obj_span != null) {
				obj_span.style.display = 'none';
			}
		}
		else if (document.frm_pesquisa.tipo_campo1.value == 'TABELA') 
		{
			if (document.frm_pesquisa.cb_campo1 != null) 
			{
				document.frm_pesquisa.cb_campo1.value = 0;
			}
		} 		
	}
	if (document.frm_pesquisa.tipo_campo2 != null) {
		if (document.frm_pesquisa.tipo_campo2.value == 'ALFA') {
			if (document.frm_pesquisa.campo2 != null) {
				document.frm_pesquisa.campo2.value = '';
			}
		} else if (document.frm_pesquisa.tipo_campo2.value == 'LOGICO') {
			if (document.frm_pesquisa.campo2 != null) {
				document.frm_pesquisa.campo2.value = '';
			}
		} else if (document.frm_pesquisa.tipo_campo2.value == 'NUM') {
			if (document.frm_pesquisa.cb_campo2 != null) {
				document.frm_pesquisa.cb_campo2.value = 0;
			}
			if (document.frm_pesquisa.campo2_ini != null) {
				document.frm_pesquisa.campo2_ini.value = '';
			}
			if (document.frm_pesquisa.campo2_fim != null) {
				document.frm_pesquisa.campo2_fim.value = '';
			}
			var obj_span = MM_findObj('span_campo2');
			if (obj_span != null) {
				obj_span.style.display = 'none';
			}
		} else if (document.frm_pesquisa.tipo_campo2.value == 'DATA') {
			if (document.frm_pesquisa.cb_campo2 != null) {
				document.frm_pesquisa.cb_campo2.value = 0;
			}
			if (document.frm_pesquisa.campo2_ini_d != null) {
				document.frm_pesquisa.campo2_ini_d.value = '';
			}
			if (document.frm_pesquisa.campo2_ini_m != null) {
				document.frm_pesquisa.campo2_ini_m.value = '';
			}
			if (document.frm_pesquisa.campo2_ini_a != null) {
				document.frm_pesquisa.campo2_ini_a.value = '';
			}
			if (document.frm_pesquisa.campo2_fim_d != null) {
				document.frm_pesquisa.campo2_fim_d.value = '';
			}
			if (document.frm_pesquisa.campo2_fim_m != null) {
				document.frm_pesquisa.campo2_fim_m.value = '';
			}
			if (document.frm_pesquisa.campo2_fim_a != null) {
				document.frm_pesquisa.campo2_fim_a.value = '';
			}
			var obj_span = MM_findObj('span_campo2');
			if (obj_span != null) {
				obj_span.style.display = 'none';
			}
		}
		else if (document.frm_pesquisa.tipo_campo2.value == 'TABELA') 
		{
			if (document.frm_pesquisa.cb_campo2 != null) 
			{
				document.frm_pesquisa.cb_campo2.value = 0;
			}
		} 		
	}
	if (document.frm_pesquisa.tipo_campo3 != null) {
		if (document.frm_pesquisa.tipo_campo3.value == 'ALFA') {
			if (document.frm_pesquisa.campo3 != null) {
				document.frm_pesquisa.campo3.value = '';
			}
		} else if (document.frm_pesquisa.tipo_campo3.value == 'LOGICO') {
			if (document.frm_pesquisa.campo3 != null) {
				document.frm_pesquisa.campo3.value = '';
			}
		} else if (document.frm_pesquisa.tipo_campo3.value == 'NUM') {
			if (document.frm_pesquisa.cb_campo3 != null) {
				document.frm_pesquisa.cb_campo3.value = 0;
			}
			if (document.frm_pesquisa.campo3_ini != null) {
				document.frm_pesquisa.campo3_ini.value = '';
			}
			if (document.frm_pesquisa.campo3_fim != null) {
				document.frm_pesquisa.campo3_fim.value = '';
			}
			var obj_span = MM_findObj('span_campo3');
			if (obj_span != null) {
				obj_span.style.display = 'none';
			}
		} else if (document.frm_pesquisa.tipo_campo3.value == 'DATA') {
			if (document.frm_pesquisa.cb_campo3 != null) {
				document.frm_pesquisa.cb_campo3.value = 0;
			}
			if (document.frm_pesquisa.campo3_ini_d != null) {
				document.frm_pesquisa.campo3_ini_d.value = '';
			}
			if (document.frm_pesquisa.campo3_ini_m != null) {
				document.frm_pesquisa.campo3_ini_m.value = '';
			}
			if (document.frm_pesquisa.campo3_ini_a != null) {
				document.frm_pesquisa.campo3_ini_a.value = '';
			}
			if (document.frm_pesquisa.campo3_fim_d != null) {
				document.frm_pesquisa.campo3_fim_d.value = '';
			}
			if (document.frm_pesquisa.campo3_fim_m != null) {
				document.frm_pesquisa.campo3_fim_m.value = '';
			}
			if (document.frm_pesquisa.campo3_fim_a != null) {
				document.frm_pesquisa.campo3_fim_a.value = '';
			}
			var obj_span = MM_findObj('span_campo3');
			if (obj_span != null) {
				obj_span.style.display = 'none';
			}
		}
		else if (document.frm_pesquisa.tipo_campo3.value == 'TABELA') 
		{
			if (document.frm_pesquisa.cb_campo3 != null) 
			{
				document.frm_pesquisa.cb_campo3.value = 0;
			}
		} 		
	}
	if (document.frm_pesquisa.tipo_campo4 != null) {
		if (document.frm_pesquisa.tipo_campo4.value == 'ALFA') {
			if (document.frm_pesquisa.campo4 != null) {
				document.frm_pesquisa.campo4.value = '';
			}
		} else if (document.frm_pesquisa.tipo_campo4.value == 'LOGICO') {
			if (document.frm_pesquisa.campo4 != null) {
				document.frm_pesquisa.campo4.value = '';
			}
		} else if (document.frm_pesquisa.tipo_campo4.value == 'NUM') {
			if (document.frm_pesquisa.cb_campo4 != null) {
				document.frm_pesquisa.cb_campo4.value = 0;
			}
			if (document.frm_pesquisa.campo4_ini != null) {
				document.frm_pesquisa.campo4_ini.value = '';
			}
			if (document.frm_pesquisa.campo4_fim != null) {
				document.frm_pesquisa.campo4_fim.value = '';
			}
			var obj_span = MM_findObj('span_campo4');
			if (obj_span != null) {
				obj_span.style.display = 'none';
			}
		} else if (document.frm_pesquisa.tipo_campo4.value == 'DATA') {
			if (document.frm_pesquisa.cb_campo4 != null) {
				document.frm_pesquisa.cb_campo4.value = 0;
			}
			if (document.frm_pesquisa.campo4_ini_d != null) {
				document.frm_pesquisa.campo4_ini_d.value = '';
			}
			if (document.frm_pesquisa.campo4_ini_m != null) {
				document.frm_pesquisa.campo4_ini_m.value = '';
			}
			if (document.frm_pesquisa.campo4_ini_a != null) {
				document.frm_pesquisa.campo4_ini_a.value = '';
			}
			if (document.frm_pesquisa.campo4_fim_d != null) {
				document.frm_pesquisa.campo4_fim_d.value = '';
			}
			if (document.frm_pesquisa.campo4_fim_m != null) {
				document.frm_pesquisa.campo4_fim_m.value = '';
			}
			if (document.frm_pesquisa.campo4_fim_a != null) {
				document.frm_pesquisa.campo4_fim_a.value = '';
			}
			var obj_span = MM_findObj('span_campo4');
			if (obj_span != null) {
				obj_span.style.display = 'none';
			}
		}
		else if (document.frm_pesquisa.tipo_campo4.value == 'TABELA') 
		{
			if (document.frm_pesquisa.cb_campo4 != null) 
			{
				document.frm_pesquisa.cb_campo4.value = 0;
			}
		} 		
	}
	if (edDados != null) {
		edDados.focus();
	}
}

function habilitaEntre(tipo,id) {
	//CAMPO NUMERICO
	if (tipo == 'num') {
		var data = MM_findObj('cb_campo' + id);
		if (data != null) {
			//Entre
			if (data.value == 3) {
				var cmp = MM_findObj('span_campo' + id);
				if (cmp != null) {
					cmp.style.display = "inline";
				}
				cmp = MM_findObj('campo' + id + '_ini');
				if (cmp != null) {
					cmp.focus();
				}
			//Igual, Maior que, Menor que
			} else {
				var cmp = MM_findObj('span_campo' + id);
				if (cmp != null) {
					cmp.style.display = "none";
				}
				cmp = MM_findObj('campo' + id + '_fim');
				if (cmp != null) {
					cmp.value = '';
				}
				cmp = MM_findObj('campo' + id + '_ini');
				if (cmp != null) {
					cmp.focus();
				}
			}
		}
	//CAMPO DATA
	} else {
		var data = MM_findObj('cb_campo' + id);
		if (data != null) {
			//Entre
			if (data.value == 3) {
				var cmp = MM_findObj('span_campo' + id);
				if (cmp != null) {
					cmp.style.display = "inline";
				}
				cmp = MM_findObj('campo' + id + '_ini_d');
				if (cmp != null) {
					cmp.focus();
				}
			//Igual, Maior que, Menor que
			} else {
				var cmp = MM_findObj('span_campo' + id);
				if (cmp != null) {
					cmp.style.display = "none";
				}
				cmp = MM_findObj('campo' + id + '_fim_d');
				if (cmp != null) {
					cmp.value = '';
				}
				cmp = MM_findObj('campo' + id + '_fim_m');
				if (cmp != null) {
					cmp.value = '';
				}
				cmp = MM_findObj('campo' + id + '_fim_a');
				if (cmp != null) {
					cmp.value = '';
				}
				cmp = MM_findObj('campo' + id + '_ini_d');
				if (cmp != null) {
					cmp.focus();
				}
			}
		}
	}
}

//**************************
// ROTINAS GLOBAIS
//**************************

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function LinkVoltar(content,paramURL) {
	if (content == "pesquisa") {
		ajxStartLoad('asp/pesquisa.asp?content='+content+'&veio_de=voltar'+paramURL);
	} else if (content == "detalhe_resultados") {
		ajxStartLoad('asp/resumo.asp?content='+content+'&veio_de=voltar'+paramURL);
	} else if (content == "resultados") {
		ajxStartLoad('asp/resultado.asp?content='+content+'&veio_de=voltar'+paramURL);
	} else if (content == "detalhe") {
		ajxStartLoad('asp/detalhes.asp?content='+content+'&veio_de=voltar'+paramURL);
	} else {
		document.location = 'index.asp?content='+content+'&veio_de=voltar'+paramURL;
	}
}

function NovaPesquisa() {
	ajxStartLoad('asp/pesquisa.asp?ajax=1');
}

//**************************
// ROTINAS RESULTADO BUSCA
//**************************

function exibeListagem(sTmp,iTmpMaterial,iTmpContexto,sDados,iObjeto,iContexto,sCampo1,sCampo2,sCampo3,sCampo4) {
	var url = "asp/resultado.asp?content=resultados&dados=" + sDados + 
			  "&objeto=" + iObjeto + "&contexto=" + iContexto + "&campo1=" + sCampo1 +
			  "&campo2=" + sCampo2 + "&campo3=" + sCampo3 +"&campo4=" + sCampo4 +
			  "&tmp=" + sTmp + "&tmp_objeto=" + iTmpMaterial + "&tmp_contexto=" + iTmpContexto;
	ajxStartLoad(url);
}

//**************************
// ROTINAS DETALHE
//**************************

function imgPreferencial(codItem,codImg) {
	var url = "asp/zoom.asp?item=" + codItem + "&imagem=" + codImg + "&zoom=0";
	window.open(url,'ImagemPopup','location=no,scrollbars=yes,menubars=no,toolbars=no,resizable=yes,left=25,top=25,width=400,height=400');
}

function LinkImagem(codItem) {
	var iAtual = parent.hiddenFrame.img_atual;
	var url    = "asp/zoom.asp?item=" + codItem + "&imagem=" + parent.hiddenFrame.img_codigos[iAtual] + 
				 "&zoom=1&content=" + parent.hiddenFrame.img_content[iAtual];
	window.open(url,'ImagemPopup','location=no,scrollbars=yes,menubars=no,toolbars=no,resizable=yes,left=25,top=25,width=750,height=550');
}

function LinkMidia(codMidia,ext,arq) {
	var url = "asp/midia.asp?codigo=" + codMidia + "&ext=" + ext + "&arq=" + arq;
	window.open(url,'MidiaPopup','location=yes,scrollbars=yes,menubars=yes,toolbars=yes,resizable=yes,left=25,top=25,width=750,height=550');
}

function LinkProx(codItem) {
	if (parent.hiddenFrame.img_total > 0) {
		if (parent.hiddenFrame.img_atual < (parent.hiddenFrame.img_total-1)) {
			var iProx = parent.hiddenFrame.img_atual + 1;
		} else {
			var iProx = 0;
		}
		var url = "asp/imagem.asp?item=" + codItem + "&imagem=" + parent.hiddenFrame.img_codigos[iProx] + "&zoom=0";
		var img = document.getElementById("img_item");
		if (img != null) {
			img.src = url;
			parent.hiddenFrame.img_atual = iProx;
			var iNumImg = document.getElementById("iNumImg");
			if (iNumImg != null) {
				iNumImg.innerHTML = "<br><b>" + (iProx+1) + "</b>&nbsp;/&nbsp;<b>" + parent.hiddenFrame.img_total + "</b>";
			}
		}
	}
}

function LinkAnt(codItem) {
	if (parent.hiddenFrame.img_total > 0) {
		if (parent.hiddenFrame.img_atual > 0) {
			var iProx = parent.hiddenFrame.img_atual - 1;
		} else {
			var iProx = (parent.hiddenFrame.img_total-1);
		}
		var url = "asp/imagem.asp?item=" + codItem + "&imagem=" + parent.hiddenFrame.img_codigos[iProx] + "&zoom=0";
		var img = document.getElementById("img_item");
		if (img != null) {
			img.src = url;
			parent.hiddenFrame.img_atual = iProx;
			var iNumImg = document.getElementById("iNumImg");
			if (iNumImg != null) {
				iNumImg.innerHTML = "<br><b>" + (iProx+1) + "</b>&nbsp;/&nbsp;<b>" + parent.hiddenFrame.img_total + "</b>";
			}
		}
	}
}