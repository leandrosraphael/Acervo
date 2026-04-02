/*****************************************************************/
/****************** ATUALIZA O NOME DO USUARIO *******************/
/******* A REQUISIÇÃO E A ATUALIZAÇÃO É FEITA COM AJAX ***********/
/*****************************************************************/

var http_request_solic = false;
var bSubmit = false;

function validaCodigoUsuSolic(event,servidor,idioma) {
	if(navigator.appName.indexOf("Netscape")!= -1) {
		tecla= event.which; 
	} else {
		tecla= event.keyCode; 
	}
	if ( tecla == 13 ) {
		bSubmit = false;
		atualizaNome(servidor,idioma);
	}
	return true;
}

function enviaSolic(servidor,idioma) {
	if ((document.frmListaImp.CodUsu != null)&&(document.frmListaImp.NomeUsu != null)&&(document.frmListaImp.Codigo != null)) {
		if (document.frmListaImp.CodUsu.value != '') {
			if (document.frmListaImp.CodUsu_ != null) {
				document.frmListaImp.submit();
			} else {
				bSubmit = true;
				atualizaNome(servidor,idioma);
			}
		} else {
			if (document.frmListaImp.NomeUsu.value == '') {
				alert(getTermo(idioma, 474, 'O nome deve ser preenchido.', 0));
				document.frmListaImp.NomeUsu.focus();
			} else {
				document.frmListaImp.submit();
			}
		}
	}
}

function atualizaNome(servidor,idioma) {
	//NOME
	if ((document.frmListaImp.CodUsu != null)&&(document.frmListaImp.NomeUsu != null)&&(document.frmListaImp.Codigo != null)) {
		if (document.frmListaImp.CodUsu.value != '') {
			var codigo = document.frmListaImp.CodUsu.value;
			requestSolic(servidor,idioma,codigo);
		}
	}
}

function requestSolic(servidor,idioma,codigo) {

	http_request_solic = false;

	if (window.XMLHttpRequest) { // Mozilla, Safari,...
		http_request_solic = new XMLHttpRequest();
		if (http_request_solic.overrideMimeType) {
			http_request_solic.overrideMimeType('text/xml');
			// See note below about this line
		}
	} else if (window.ActiveXObject) { // IE
		try {
			http_request_solic = new ActiveXObject("Msxml2.XMLHTTP");
		} catch (e) {
			try {
				http_request_solic = new ActiveXObject("Microsoft.XMLHTTP");
			} catch (e) {}
		}
	}

	if (!http_request_solic) {
		alert(getTermo(idioma, 1270, 'Ocorreu um erro interno.', 0));
		return false;
	}
	
	var url = 'usuario_xml.asp?codigo='+codigo+'&servidor='+servidor+'&iIdioma='+idioma;
	
	http_request_solic.onreadystatechange = processaSolic;
	http_request_solic.open('GET', url, true);
	http_request_solic.send(null);
}

function processaSolic() {
	var idioma = 0;
	if (parent.parent.hiddenFrame != null) {
		idioma = parent.parent.hiddenFrame.iIdioma;
	}
	
	if (http_request_solic.readyState == 4) {
		if (http_request_solic.status == 200) {
			var xmldoc = http_request_solic.responseXML;
			
			//NOME DO USUARIO
			if (document.frmListaImp.NomeUsu != null) {
				var xNome = xmldoc.getElementsByTagName('NOME');
				//Se possuir series para serem exibidos inclui no combo
				if (xNome.length > 0) {
					xNome = xNome[0];
					var DescUsu = xNome.getAttribute('VALOR');
					if (DescUsu == '') {
						alert(getTermo(idioma, 1206, 'Usuário não existe.', 0));
						document.frmListaImp.Codigo.value = '';
						document.frmListaImp.CodUsu.focus();
						document.frmListaImp.CodUsu.select();
					} else {
						document.frmListaImp.Codigo.value = '';
						document.frmListaImp.NomeUsu.value = DescUsu;
						document.frmListaImp.NomeUsu.focus();
						document.frmListaImp.NomeUsu.select();
					}
					var xCodigo = xmldoc.getElementsByTagName('CODIGO');
					if (xCodigo.length > 0) {
						xCodigo = xCodigo[0];
						var CodUsu = xCodigo.getAttribute('VALOR');
						if (CodUsu != '') {
							document.frmListaImp.Codigo.value = CodUsu;
						}
					}
				} else {
					document.frmListaImp.Codigo.value = '';
					alert(getTermo(idioma, 1206, 'Usuário não existe.', 0));
					document.frmListaImp.CodUsu.focus();
					document.frmListaImp.CodUsu.select();
				}			
			}
			
			if (bSubmit) {
				//faz o submit
				if (! ((document.frmListaImp.CodUsu.value != '') && (document.frmListaImp.Codigo.value == ''))) {
					if (document.frmListaImp.NomeUsu.value == '') {
						alert(getTermo(idioma, 474, 'O nome deve ser preenchido.', 0));
						document.frmListaImp.NomeUsu.focus();
					} else {
						document.frmListaImp.submit();
					}
				}
			}
		} else {
			alert(getTermo(idioma, 1271, 'Ocorreu um erro durante a conexão com o servidor.', 0));
		}
	}
}