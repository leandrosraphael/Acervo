//**************************
// ROTINAS AJAX
//**************************

function posTop() {
	var win = self;
	var doc = document;
	return typeof win.pageYOffset != 'undefined' ? win.pageYOffset:doc.documentElement && doc.documentElement.scrollTop? doc.documentElement.scrollTop: doc.body.scrollTop?doc.body.scrollTop:0;
}

function startLoading() {
	div = document.getElementById('dvLoad');
	if (div != null) {
		div.style.top = posTop() + 150;
		div.style.display = 'block';
	}
}

function finishLoading() {
    div = document.getElementById('dvLoad');
    if (div != null) {
        div.style.display = 'none';
    }
}

function ajxStartLoad(URL, callback) {
	startLoading();

	if (URL != "") {
		$.ajax({
			url: URL,
			cache: false,
			success: function (data, textStatus, jqXHR) {
			    ajxFinishLoad(data, callback);
			}
		});
	}
}

function ajxFormPesquisaSubmit() {
	var url = $('#frm_pesquisa').attr('action');
	var data = $('#frm_pesquisa').serialize();

	startLoading();

	$.ajax({
		url: url,
		cache: false,
		type: 'POST',
		data: data,
		success: function (data, textStatus, jqXHR) {
			ajxFinishLoad(data);
		}
	});
}

function ajxFinishLoad(HTML, callback) {
	div = document.getElementById('ajxDiv');
	if (div != null) {
		div.innerHTML = HTML;

		var scriptTag = document.getElementById('script');
		if (scriptTag != null && scriptTag != undefined) {
		    var s = scriptTag.innerHTML;
			eval(s);
		}
	}

	if (callback) {
	    callback();
	}

	finishLoading();
	$('#dados_outro').focus();
	window.scroll(0, 0);
}

//**************************
// ROTINAS VALIDAÇÃO BUSCA
//**************************

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
	    AtualizarComboMaterialOrdenacao(0);
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
	if (content == "detalhe_resultados") {
	    ajxStartLoad('asp/resumo.asp?content=' + content + '&veio_de=voltar' + paramURL);
	} else if (content == "resultados") {
	    ajxStartLoad('asp/resultado.asp?content=' + content + '&veio_de=voltar' + paramURL);
	} else if (content == "detalhe") {
	    ajxStartLoad('asp/detalhes.asp?content=' + content + '&veio_de=voltar' + paramURL);
	} else {
		document.location = 'index.asp?content='+content+'&veio_de=voltar'+paramURL;
	}
}

function LinkVoltarPesquisa(paramURL) {

    var callback = function () {
        var idSelect = $("#objeto option:selected").val();
        AtualizarComboMaterialOrdenacao(idSelect, paramURL);
    };

    ajxStartLoad('asp/pesquisa.asp' + paramURL, callback);
}

function NovaPesquisa() {
	ajxStartLoad('asp/pesquisa.asp?ajax=1', AtualizarComboMaterialOrdenacao(0));
}

//**************************
// ROTINAS RESULTADO BUSCA
//**************************

function exibeListagem(sTmp, iTmpMaterial, iTmpContexto, sDados, iObjeto, iContexto, sCampo1, sCampo2, sCampo3, sCampo4, sCampo5, sCampo6, sCampo7, sCampo8, campoOrdenacao, campoOrdem) {
	var url = "asp/resultado.asp?content=resultados&dados=" + sDados +
		"&objeto=" + iObjeto + "&contexto=" + iContexto +
		"&campo1=" + sCampo1 + "&campo2=" + sCampo2 + "&campo3=" + sCampo3 + "&campo4=" + sCampo4 +
		"&campo5=" + sCampo5 + "&campo6=" + sCampo6 + "&campo7=" + sCampo7 + "&campo8=" + sCampo8 +
		"&tmp=" + sTmp + "&tmp_objeto=" + iTmpMaterial + "&tmp_contexto=" + iTmpContexto +
		"&campo_ordenacao=" + campoOrdenacao + "&campo_ordem=" + campoOrdem;

	ajxStartLoad(url);
}

var getUrlParameter = function getUrlParameter(sParam, url) {
	var sPageURL = decodeURIComponent(url),
		sURLVariables = sPageURL.split('&'),
		sParameterName,
		i;

	for (i = 0; i < sURLVariables.length; i++) {
		sParameterName = sURLVariables[i].split('=');

		if (sParameterName[0] === sParam) {
			return sParameterName[1] === undefined ? true : sParameterName[1];
		}
	}
};

function carregaPagina(paramURL, pagina) {
	if (pagina) {
		var paginaParam = getUrlParameter('pagina', paramURL);
		paramURL = paramURL.replace('&pagina=' + paginaParam, '&pagina=' + pagina);
	}

	var url = "asp/resultado.asp?content=resultados" + paramURL;
	ajxStartLoad(url);
}

//**************************
// ROTINAS DETALHE
//**************************

function LinkImagem(codItem) {
	var iAtual = parent.hiddenFrame.img_atual;
	var url    = "asp/zoom.asp?item=" + codItem + "&imagem=" + parent.hiddenFrame.img_codigos[iAtual] + 
				 "&zoom=1&content=" + parent.hiddenFrame.img_content[iAtual];
	window.open(url,'ImagemPopup','location=no,scrollbars=yes,menubars=no,toolbars=no,resizable=yes,left=25,top=25,width=750,height=550');
}

function LinkMidia(codMidia,ext,arq) {
    var url = "asp/midia.asp?codigo=" + codMidia + "&ext=" + ext + "&arq=" + arq;
    window.open(url, 'MidiaPopup', 'location=yes,scrollbars=yes,menubars=yes,toolbars=yes,resizable=yes,left=25,top=25,width=750,height=550');
}

function LinkMidiaPdf(codMidia, arq) {
    var url = "asp/prima-pdf.asp?codigoMidia=" + codMidia + "&nomeArquivo=" + arq;
    window.open(url);
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

function AbrirTelaLogin() {
    div = $('#DivLogin');
    if (div != null) {
        div.css('display', 'block');
        $('#screen').css({ "display": "block", opacity: 0.4, "width": $(document).width(), "height": $(document).height() });
        $("#DivLogin").load('asp/login.asp'); 
    }
}

function FecharTelaLogin() {
    div = $('#DivLogin');
    if (div != null) {
        $('#screen').css({ "display": "none" });
        div.css('display', 'none');
    }
}

function EfetuarLogin() {
    var senha = document.login.senha.value;
    var login = document.login.loginUsuario.value;
    var labelLogin = document.login.LabelLogin.value;
    if (login.length == 0 || senha.length == 0) {
        if (login.length == 0) {
            if (loginPorMatricula) {
                var mensagem = "O campo " + labelLogin + " não pode ser vazio."
            } else {
                var mensagem = "O campo Código não pode ser vazio.";
            }
            document.login.loginUsuario.focus();
        } else if (senha.length == 0) {
            var mensagem = "O campo Senha não pode ser vazio.";
            document.login.senha.focus();
        }
        alert(mensagem);
        return false;
    }

    $.ajax({
    	url: 'asp/efetuarLogin.asp?login=' + login + '&senha=' + senha,
    	cache: false
    }).done(function (data) {
        if (data.logou == "true") {
            FecharTelaLogin();
            $("#a-login").text('Sair');
            $("#a-login").attr("onclick", "EfetuarLogout()");
            $("#a-nome-login").text(data.nome + ',');
            location.reload();
        } else {
            alert(data.mensagem);
        }
    });
}

function EfetuarLogout() {
    $.ajax({
    	url: 'asp/efetuarLogout.asp',
    	cache: false
    }).done(function () {
        $("#a-login").text('Login');
        $("#a-login").attr("onclick", "AbrirTelaLogin()");
        $("#a-nome-login").text('');
        location.reload();
    });
}

function BloqueiaNaoNumerico(evt) {
    var charCode = (evt.which) ? evt.which : event.keyCode;
    if (charCode > 31 && (charCode < 48 || charCode > 57)) {
        return false;
    }

    return true;
}

function AtualizarComboMaterialOrdenacao(idMaterial, ParamURL) {

	var urlMat = 'asp/atualizar_combo_material.asp';
	var urlOrdenacao = 'asp/atualizar_combo_ordenacao.asp';
	var parametro = { 'idMaterial': idMaterial }

    if (ParamURL) {
		urlMat = urlMat + ParamURL;
		urlOrdenacao = urlOrdenacao + ParamURL;
    }
    
    if (idMaterial != null) {
        $.ajax({
            url: urlMat,
			cache: false,
			data: parametro,
			beforeSend: function () {
				startLoading();
			},
            success: function (data, textStatus, jqXHR) {
                $('.campo_material').remove();
                if (data) {
                    $('#combo_material').after(data);
                }
			},
			complete: function () {
				$.ajax({
					url: urlOrdenacao,
					data: parametro,
					cache: false,
					success: function (data, textStatus, jqXHR) {
						$('#campos_ordenacao').empty().append(data);
					}
				});
			}
		}).always(function () {
			finishLoading();
		});
    }
}

function ValidarData(Dia, Mes, Ano) {
    var msg = "";

    if ((dia > 31) || (dia == '')) {
        Msg = 'Digite um dia válido!';
    } else if ((mes > 12) || (mes == '')) {
        Msg = 'Digite um mês válido!';
    } else if (ano == '') {
        Msg = 'Digite um ano válido!';
    }

    if (Msg != "") {
        alert(Msg);
        return false;
    } else {
        return true;
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
            ajxFormPesquisaSubmit();
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
            else if (document.frm_pesquisa.tipo_campo1.value == 'TABELA') {
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
                    if (ValidaCamposDataVazio(document.frm_pesquisa.campo1_ini_d.value, document.frm_pesquisa.campo1_ini_m.value, document.frm_pesquisa.campo1_ini_a.value,
                         document.frm_pesquisa.campo1_fim_d.value, document.frm_pesquisa.campo1_fim_m.value, document.frm_pesquisa.campo1_fim_a.value)) {

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

                        if ((document.frm_pesquisa.campo1_ini_d.value == '') || (document.frm_pesquisa.campo1_ini_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo1_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo1_ini_m.value == '') || (document.frm_pesquisa.campo1_ini_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo1_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo1_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo1_ini_a.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo1_fim_d.value == '') || (document.frm_pesquisa.campo1_fim_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo1_fim_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo1_fim_m.value == '') || (document.frm_pesquisa.campo1_fim_m.value > 12)) {
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
            else if (document.frm_pesquisa.tipo_campo2.value == 'TABELA') {
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
                    if (ValidaCamposDataVazio(document.frm_pesquisa.campo2_ini_d.value, document.frm_pesquisa.campo2_ini_m.value, document.frm_pesquisa.campo2_ini_a.value,
                         document.frm_pesquisa.campo2_fim_d.value, document.frm_pesquisa.campo2_fim_m.value, document.frm_pesquisa.campo2_fim_a.value)) {

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

                        if ((document.frm_pesquisa.campo2_ini_d.value == '') || (document.frm_pesquisa.campo2_ini_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo2_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo2_ini_m.value == '') || (document.frm_pesquisa.campo2_ini_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo2_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo2_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo2_ini_a.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo2_fim_d.value == '') || (document.frm_pesquisa.campo2_fim_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo2_fim_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo2_fim_m.value == '') || (document.frm_pesquisa.campo2_fim_m.value > 12)) {
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
            else if (document.frm_pesquisa.tipo_campo3.value == 'TABELA') {
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
                    if (ValidaCamposDataVazio(document.frm_pesquisa.campo3_ini_d.value, document.frm_pesquisa.campo3_ini_m.value, document.frm_pesquisa.campo3_ini_a.value,
                         document.frm_pesquisa.campo3_fim_d.value, document.frm_pesquisa.campo3_fim_m.value, document.frm_pesquisa.campo3_fim_a.value)) {

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

                        if ((document.frm_pesquisa.campo3_ini_d.value == '') || (document.frm_pesquisa.campo3_ini_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo3_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo3_ini_m.value == '') || (document.frm_pesquisa.campo3_ini_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo3_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo3_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo3_ini_a.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo3_fim_d.value == '') || (document.frm_pesquisa.campo3_fim_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo3_fim_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo3_fim_m.value == '') || (document.frm_pesquisa.campo3_fim_m.value > 12)) {
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
            else if (document.frm_pesquisa.tipo_campo4.value == 'TABELA') {
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
                    if (ValidaCamposDataVazio(document.frm_pesquisa.campo4_ini_d.value, document.frm_pesquisa.campo4_ini_m.value, document.frm_pesquisa.campo4_ini_a.value,
                         document.frm_pesquisa.campo4_fim_d.value, document.frm_pesquisa.campo4_fim_m.value, document.frm_pesquisa.campo4_fim_a.value)) {

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

                        if ((document.frm_pesquisa.campo4_ini_d.value == '') || (document.frm_pesquisa.campo4_ini_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo4_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo4_ini_m.value == '') || (document.frm_pesquisa.campo4_ini_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo4_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo4_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo4_ini_a.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo4_fim_d.value == '') || (document.frm_pesquisa.campo4_fim_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo4_fim_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo4_fim_m.value == '') || (document.frm_pesquisa.campo4_fim_m.value > 12)) {
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
        //VALIDA CAMPO OPCIONAL 5
        if (document.frm_pesquisa.tipo_campo5 != null) {
            if (document.frm_pesquisa.tipo_campo5.value == 'ALFA') {
                if (document.frm_pesquisa.campo5 != null) {
                    tam_dados += document.frm_pesquisa.campo5.value.length;
                }
            } else if (document.frm_pesquisa.tipo_campo5.value == 'LOGICO') {
                if (document.frm_pesquisa.campo5 != null) {
                    if (document.frm_pesquisa.campo5.value != '') {
                        tam_dados += 1;
                    }
                }
            } else if (document.frm_pesquisa.tipo_campo5.value == 'NUM') {
                if (document.frm_pesquisa.cb_campo5.value == '3') {
                    if (document.frm_pesquisa.campo5_ini.value == '') {
                        alert('Preencha o intervalo inicial!');
                        document.frm_pesquisa.campo5_ini.focus();
                        return false;
                    } else if (document.frm_pesquisa.campo5_fim.value == '') {
                        alert('Preencha o intervalo final!');
                        document.frm_pesquisa.campo5_fim.focus();
                        return false;
                    } else {
                        tam_dados += 1;
                    }
                } else {
                    if (document.frm_pesquisa.campo5_ini.value != '') {
                        tam_dados += 1;
                    }
                }
            }
            else if (document.frm_pesquisa.tipo_campo5.value == 'TABELA') {
                tam_dados += document.frm_pesquisa.cb_campo5.value;
            }
            else if (document.frm_pesquisa.tipo_campo5.value == 'DATA') {
                if (document.frm_pesquisa.campo5_ini_d.value.length == 1) {
                    document.frm_pesquisa.campo5_ini_d.value = '0' + document.frm_pesquisa.campo5_ini_d.value;
                }
                if (document.frm_pesquisa.campo5_ini_m.value.length == 1) {
                    document.frm_pesquisa.campo5_ini_m.value = '0' + document.frm_pesquisa.campo5_ini_m.value;
                }
                if (document.frm_pesquisa.campo5_ini_a.value.length == 1) {
                    document.frm_pesquisa.campo5_ini_a.value = '200' + document.frm_pesquisa.campo5_ini_a.value;
                } else if (document.frm_pesquisa.campo5_ini_a.value.length == 2) {
                    if (document.frm_pesquisa.campo5_ini_a.value > 50) {
                        document.frm_pesquisa.campo5_ini_a.value = '20' + document.frm_pesquisa.campo5_ini_a.value;
                    } else {
                        document.frm_pesquisa.campo5_ini_a.value = '19' + document.frm_pesquisa.campo5_ini_a.value;
                    }
                } else if (document.frm_pesquisa.campo5_ini_a.value.length == 3) {
                    document.frm_pesquisa.campo5_ini_a.value = '0' + document.frm_pesquisa.campo5_ini_a.value;
                }

                if (document.frm_pesquisa.cb_campo5.value == '3') {
                    if (ValidaCamposDataVazio(document.frm_pesquisa.campo5_ini_d.value, document.frm_pesquisa.campo5_ini_m.value, document.frm_pesquisa.campo5_ini_a.value,
                         document.frm_pesquisa.campo5_fim_d.value, document.frm_pesquisa.campo5_fim_m.value, document.frm_pesquisa.campo5_fim_a.value)) {

                        if (document.frm_pesquisa.campo5_fim_d.value.length == 1) {
                            document.frm_pesquisa.campo5_fim_d.value = '0' + document.frm_pesquisa.campo5_fim_d.value;
                        }
                        if (document.frm_pesquisa.campo5_fim_m.value.length == 1) {
                            document.frm_pesquisa.campo5_fim_m.value = '0' + document.frm_pesquisa.campo5_fim_m.value;
                        }
                        if (document.frm_pesquisa.campo5_fim_a.value.length == 1) {
                            document.frm_pesquisa.campo5_fim_a.value = '200' + document.frm_pesquisa.campo5_fim_a.value;
                        } else if (document.frm_pesquisa.campo5_fim_a.value.length == 2) {
                            if (document.frm_pesquisa.campo5_fim_a.value > 50) {
                                document.frm_pesquisa.campo5_fim_a.value = '20' + document.frm_pesquisa.campo5_fim_a.value;
                            } else {
                                document.frm_pesquisa.campo5_fim_a.value = '19' + document.frm_pesquisa.campo5_fim_a.value;
                            }
                        } else if (document.frm_pesquisa.campo5_fim_a.value.length == 3) {
                            document.frm_pesquisa.campo5_fim_a.value = '0' + document.frm_pesquisa.campo5_fim_a.value;
                        }

                        if ((document.frm_pesquisa.campo5_ini_d.value == '') || (document.frm_pesquisa.campo5_ini_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo5_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo5_ini_m.value == '') || (document.frm_pesquisa.campo5_ini_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo5_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo5_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo5_ini_a.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo5_fim_d.value == '') || (document.frm_pesquisa.campo5_fim_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo5_fim_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo5_fim_m.value == '') || (document.frm_pesquisa.campo5_fim_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo5_fim_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo5_fim_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo5_fim_a.focus();
                            return false;
                        } else {
                            tam_dados += 1;
                        }
                    }
                } else {
                    if ((document.frm_pesquisa.campo5_ini_d.value != '') || (document.frm_pesquisa.campo5_ini_m.value != '') ||
						(document.frm_pesquisa.campo5_ini_a.value != '')) {
                        if ((document.frm_pesquisa.campo5_ini_d.value > 31) || (document.frm_pesquisa.campo5_ini_d.value == '')) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo5_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo5_ini_m.value > 12) || (document.frm_pesquisa.campo5_ini_m.value == '')) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo5_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo5_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo5_ini_a.focus();
                            return false;
                        } else {
                            tam_dados += 1;
                        }
                    }
                }
            }
        }

        //VALIDA CAMPO OPCIONAL 6
        if (document.frm_pesquisa.tipo_campo6 != null) {
            if (document.frm_pesquisa.tipo_campo6.value == 'ALFA') {
                if (document.frm_pesquisa.campo6 != null) {
                    tam_dados += document.frm_pesquisa.campo6.value.length;
                }
            } else if (document.frm_pesquisa.tipo_campo6.value == 'LOGICO') {
                if (document.frm_pesquisa.campo6 != null) {
                    if (document.frm_pesquisa.campo6.value != '') {
                        tam_dados += 1;
                    }
                }
            } else if (document.frm_pesquisa.tipo_campo6.value == 'NUM') {
                if (document.frm_pesquisa.cb_campo6.value == '3') {
                    if (document.frm_pesquisa.campo6_ini.value == '') {
                        alert('Preencha o intervalo inicial!');
                        document.frm_pesquisa.campo6_ini.focus();
                        return false;
                    } else if (document.frm_pesquisa.campo6_fim.value == '') {
                        alert('Preencha o intervalo final!');
                        document.frm_pesquisa.campo6_fim.focus();
                        return false;
                    } else {
                        tam_dados += 1;
                    }
                } else {
                    if (document.frm_pesquisa.campo6_ini.value != '') {
                        tam_dados += 1;
                    }
                }
            }
            else if (document.frm_pesquisa.tipo_campo6.value == 'TABELA') {
                tam_dados += document.frm_pesquisa.cb_campo6.value;
            }
            else if (document.frm_pesquisa.tipo_campo6.value == 'DATA') {
                if (document.frm_pesquisa.campo6_ini_d.value.length == 1) {
                    document.frm_pesquisa.campo6_ini_d.value = '0' + document.frm_pesquisa.campo6_ini_d.value;
                }
                if (document.frm_pesquisa.campo6_ini_m.value.length == 1) {
                    document.frm_pesquisa.campo6_ini_m.value = '0' + document.frm_pesquisa.campo6_ini_m.value;
                }
                if (document.frm_pesquisa.campo6_ini_a.value.length == 1) {
                    document.frm_pesquisa.campo6_ini_a.value = '200' + document.frm_pesquisa.campo6_ini_a.value;
                } else if (document.frm_pesquisa.campo6_ini_a.value.length == 2) {
                    if (document.frm_pesquisa.campo6_ini_a.value > 60) {
                        document.frm_pesquisa.campo6_ini_a.value = '20' + document.frm_pesquisa.campo6_ini_a.value;
                    } else {
                        document.frm_pesquisa.campo6_ini_a.value = '19' + document.frm_pesquisa.campo6_ini_a.value;
                    }
                } else if (document.frm_pesquisa.campo6_ini_a.value.length == 3) {
                    document.frm_pesquisa.campo6_ini_a.value = '0' + document.frm_pesquisa.campo6_ini_a.value;
                }

                if (document.frm_pesquisa.cb_campo6.value == '3') {
                    if (ValidaCamposDataVazio(document.frm_pesquisa.campo6_ini_d.value, document.frm_pesquisa.campo6_ini_m.value, document.frm_pesquisa.campo6_ini_a.value,
                         document.frm_pesquisa.campo6_fim_d.value, document.frm_pesquisa.campo6_fim_m.value, document.frm_pesquisa.campo6_fim_a.value)) {

                        if (document.frm_pesquisa.campo6_fim_d.value.length == 1) {
                            document.frm_pesquisa.campo6_fim_d.value = '0' + document.frm_pesquisa.campo6_fim_d.value;
                        }
                        if (document.frm_pesquisa.campo6_fim_m.value.length == 1) {
                            document.frm_pesquisa.campo6_fim_m.value = '0' + document.frm_pesquisa.campo6_fim_m.value;
                        }
                        if (document.frm_pesquisa.campo6_fim_a.value.length == 1) {
                            document.frm_pesquisa.campo6_fim_a.value = '200' + document.frm_pesquisa.campo6_fim_a.value;
                        } else if (document.frm_pesquisa.campo6_fim_a.value.length == 2) {
                            if (document.frm_pesquisa.campo6_fim_a.value > 60) {
                                document.frm_pesquisa.campo6_fim_a.value = '20' + document.frm_pesquisa.campo6_fim_a.value;
                            } else {
                                document.frm_pesquisa.campo6_fim_a.value = '19' + document.frm_pesquisa.campo6_fim_a.value;
                            }
                        } else if (document.frm_pesquisa.campo6_fim_a.value.length == 3) {
                            document.frm_pesquisa.campo6_fim_a.value = '0' + document.frm_pesquisa.campo6_fim_a.value;
                        }

                        if ((document.frm_pesquisa.campo6_ini_d.value == '') || (document.frm_pesquisa.campo6_ini_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo6_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo6_ini_m.value == '') || (document.frm_pesquisa.campo6_ini_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo6_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo6_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo6_ini_a.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo6_fim_d.value == '') || (document.frm_pesquisa.campo6_fim_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo6_fim_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo6_fim_m.value == '') || (document.frm_pesquisa.campo6_fim_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo6_fim_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo6_fim_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo6_fim_a.focus();
                            return false;
                        } else {
                            tam_dados += 1;
                        }
                    }
                } else {
                    if ((document.frm_pesquisa.campo6_ini_d.value != '') || (document.frm_pesquisa.campo6_ini_m.value != '') ||
						(document.frm_pesquisa.campo6_ini_a.value != '')) {
                        if ((document.frm_pesquisa.campo6_ini_d.value > 31) || (document.frm_pesquisa.campo6_ini_d.value == '')) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo6_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo6_ini_m.value > 12) || (document.frm_pesquisa.campo6_ini_m.value == '')) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo6_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo6_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo6_ini_a.focus();
                            return false;
                        } else {
                            tam_dados += 1;
                        }
                    }
                }
            }
        }
        //VALIDA CAMPO OPCIONAL 7
        if (document.frm_pesquisa.tipo_campo7 != null) {
            if (document.frm_pesquisa.tipo_campo7.value == 'ALFA') {
                if (document.frm_pesquisa.campo7 != null) {
                    tam_dados += document.frm_pesquisa.campo7.value.length;
                }
            } else if (document.frm_pesquisa.tipo_campo7.value == 'LOGICO') {
                if (document.frm_pesquisa.campo7 != null) {
                    if (document.frm_pesquisa.campo7.value != '') {
                        tam_dados += 1;
                    }
                }
            } else if (document.frm_pesquisa.tipo_campo7.value == 'NUM') {
                if (document.frm_pesquisa.cb_campo7.value == '3') {
                    if (document.frm_pesquisa.campo7_ini.value == '') {
                        alert('Preencha o intervalo inicial!');
                        document.frm_pesquisa.campo7_ini.focus();
                        return false;
                    } else if (document.frm_pesquisa.campo7_fim.value == '') {
                        alert('Preencha o intervalo final!');
                        document.frm_pesquisa.campo7_fim.focus();
                        return false;
                    } else {
                        tam_dados += 1;
                    }
                } else {
                    if (document.frm_pesquisa.campo7_ini.value != '') {
                        tam_dados += 1;
                    }
                }
            }
            else if (document.frm_pesquisa.tipo_campo7.value == 'TABELA') {
                tam_dados += document.frm_pesquisa.cb_campo7.value;
            }
            else if (document.frm_pesquisa.tipo_campo7.value == 'DATA') {
                if (document.frm_pesquisa.campo7_ini_d.value.length == 1) {
                    document.frm_pesquisa.campo7_ini_d.value = '0' + document.frm_pesquisa.campo7_ini_d.value;
                }
                if (document.frm_pesquisa.campo7_ini_m.value.length == 1) {
                    document.frm_pesquisa.campo7_ini_m.value = '0' + document.frm_pesquisa.campo7_ini_m.value;
                }
                if (document.frm_pesquisa.campo7_ini_a.value.length == 1) {
                    document.frm_pesquisa.campo7_ini_a.value = '200' + document.frm_pesquisa.campo7_ini_a.value;
                } else if (document.frm_pesquisa.campo7_ini_a.value.length == 2) {
                    if (document.frm_pesquisa.campo7_ini_a.value > 70) {
                        document.frm_pesquisa.campo7_ini_a.value = '20' + document.frm_pesquisa.campo7_ini_a.value;
                    } else {
                        document.frm_pesquisa.campo7_ini_a.value = '19' + document.frm_pesquisa.campo7_ini_a.value;
                    }
                } else if (document.frm_pesquisa.campo7_ini_a.value.length == 3) {
                    document.frm_pesquisa.campo7_ini_a.value = '0' + document.frm_pesquisa.campo7_ini_a.value;
                }

                if (document.frm_pesquisa.cb_campo7.value == '3') {
                    if (ValidaCamposDataVazio(document.frm_pesquisa.campo7_ini_d.value, document.frm_pesquisa.campo7_ini_m.value, document.frm_pesquisa.campo7_ini_a.value,
                         document.frm_pesquisa.campo7_fim_d.value, document.frm_pesquisa.campo7_fim_m.value, document.frm_pesquisa.campo7_fim_a.value)) {

                        if (document.frm_pesquisa.campo7_fim_d.value.length == 1) {
                            document.frm_pesquisa.campo7_fim_d.value = '0' + document.frm_pesquisa.campo7_fim_d.value;
                        }
                        if (document.frm_pesquisa.campo7_fim_m.value.length == 1) {
                            document.frm_pesquisa.campo7_fim_m.value = '0' + document.frm_pesquisa.campo7_fim_m.value;
                        }
                        if (document.frm_pesquisa.campo7_fim_a.value.length == 1) {
                            document.frm_pesquisa.campo7_fim_a.value = '200' + document.frm_pesquisa.campo7_fim_a.value;
                        } else if (document.frm_pesquisa.campo7_fim_a.value.length == 2) {
                            if (document.frm_pesquisa.campo7_fim_a.value > 70) {
                                document.frm_pesquisa.campo7_fim_a.value = '20' + document.frm_pesquisa.campo7_fim_a.value;
                            } else {
                                document.frm_pesquisa.campo7_fim_a.value = '19' + document.frm_pesquisa.campo7_fim_a.value;
                            }
                        } else if (document.frm_pesquisa.campo7_fim_a.value.length == 3) {
                            document.frm_pesquisa.campo7_fim_a.value = '0' + document.frm_pesquisa.campo7_fim_a.value;
                        }

                        if ((document.frm_pesquisa.campo7_ini_d.value == '') || (document.frm_pesquisa.campo7_ini_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo7_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo7_ini_m.value == '') || (document.frm_pesquisa.campo7_ini_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo7_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo7_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo7_ini_a.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo7_fim_d.value == '') || (document.frm_pesquisa.campo7_fim_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo7_fim_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo7_fim_m.value == '') || (document.frm_pesquisa.campo7_fim_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo7_fim_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo7_fim_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo7_fim_a.focus();
                            return false;
                        } else {
                            tam_dados += 1;
                        }
                    }
                } else {
                    if ((document.frm_pesquisa.campo7_ini_d.value != '') || (document.frm_pesquisa.campo7_ini_m.value != '') ||
						(document.frm_pesquisa.campo7_ini_a.value != '')) {
                        if ((document.frm_pesquisa.campo7_ini_d.value > 31) || (document.frm_pesquisa.campo7_ini_d.value == '')) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo7_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo7_ini_m.value > 12) || (document.frm_pesquisa.campo7_ini_m.value == '')) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo7_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo7_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo7_ini_a.focus();
                            return false;
                        } else {
                            tam_dados += 1;
                        }
                    }
                }
            }
        }
        //VALIDA CAMPO OPCIONAL 8
        if (document.frm_pesquisa.tipo_campo8 != null) {
            if (document.frm_pesquisa.tipo_campo8.value == 'ALFA') {
                if (document.frm_pesquisa.campo8 != null) {
                    tam_dados += document.frm_pesquisa.campo8.value.length;
                }
            } else if (document.frm_pesquisa.tipo_campo8.value == 'LOGICO') {
                if (document.frm_pesquisa.campo8 != null) {
                    if (document.frm_pesquisa.campo8.value != '') {
                        tam_dados += 1;
                    }
                }
            } else if (document.frm_pesquisa.tipo_campo8.value == 'NUM') {
                if (document.frm_pesquisa.cb_campo8.value == '3') {
                    if (document.frm_pesquisa.campo8_ini.value == '') {
                        alert('Preencha o intervalo inicial!');
                        document.frm_pesquisa.campo8_ini.focus();
                        return false;
                    } else if (document.frm_pesquisa.campo8_fim.value == '') {
                        alert('Preencha o intervalo final!');
                        document.frm_pesquisa.campo8_fim.focus();
                        return false;
                    } else {
                        tam_dados += 1;
                    }
                } else {
                    if (document.frm_pesquisa.campo8_ini.value != '') {
                        tam_dados += 1;
                    }
                }
            }
            else if (document.frm_pesquisa.tipo_campo8.value == 'TABELA') {
                tam_dados += document.frm_pesquisa.cb_campo8.value;
            }
            else if (document.frm_pesquisa.tipo_campo8.value == 'DATA') {
                if (document.frm_pesquisa.campo8_ini_d.value.length == 1) {
                    document.frm_pesquisa.campo8_ini_d.value = '0' + document.frm_pesquisa.campo8_ini_d.value;
                }
                if (document.frm_pesquisa.campo8_ini_m.value.length == 1) {
                    document.frm_pesquisa.campo8_ini_m.value = '0' + document.frm_pesquisa.campo8_ini_m.value;
                }
                if (document.frm_pesquisa.campo8_ini_a.value.length == 1) {
                    document.frm_pesquisa.campo8_ini_a.value = '200' + document.frm_pesquisa.campo8_ini_a.value;
                } else if (document.frm_pesquisa.campo8_ini_a.value.length == 2) {
                    if (document.frm_pesquisa.campo8_ini_a.value > 80) {
                        document.frm_pesquisa.campo8_ini_a.value = '20' + document.frm_pesquisa.campo8_ini_a.value;
                    } else {
                        document.frm_pesquisa.campo8_ini_a.value = '19' + document.frm_pesquisa.campo8_ini_a.value;
                    }
                } else if (document.frm_pesquisa.campo8_ini_a.value.length == 3) {
                    document.frm_pesquisa.campo8_ini_a.value = '0' + document.frm_pesquisa.campo8_ini_a.value;
                }

                if (document.frm_pesquisa.cb_campo8.value == '3') {
                    if (ValidaCamposDataVazio(document.frm_pesquisa.campo8_ini_d.value, document.frm_pesquisa.campo8_ini_m.value, document.frm_pesquisa.campo8_ini_a.value,
                         document.frm_pesquisa.campo8_fim_d.value, document.frm_pesquisa.campo8_fim_m.value, document.frm_pesquisa.campo8_fim_a.value)) {

                        if (document.frm_pesquisa.campo8_fim_d.value.length == 1) {
                            document.frm_pesquisa.campo8_fim_d.value = '0' + document.frm_pesquisa.campo8_fim_d.value;
                        }
                        if (document.frm_pesquisa.campo8_fim_m.value.length == 1) {
                            document.frm_pesquisa.campo8_fim_m.value = '0' + document.frm_pesquisa.campo8_fim_m.value;
                        }
                        if (document.frm_pesquisa.campo8_fim_a.value.length == 1) {
                            document.frm_pesquisa.campo8_fim_a.value = '200' + document.frm_pesquisa.campo8_fim_a.value;
                        } else if (document.frm_pesquisa.campo8_fim_a.value.length == 2) {
                            if (document.frm_pesquisa.campo8_fim_a.value > 80) {
                                document.frm_pesquisa.campo8_fim_a.value = '20' + document.frm_pesquisa.campo8_fim_a.value;
                            } else {
                                document.frm_pesquisa.campo8_fim_a.value = '19' + document.frm_pesquisa.campo8_fim_a.value;
                            }
                        } else if (document.frm_pesquisa.campo8_fim_a.value.length == 3) {
                            document.frm_pesquisa.campo8_fim_a.value = '0' + document.frm_pesquisa.campo8_fim_a.value;
                        }

                        if ((document.frm_pesquisa.campo8_ini_d.value == '') || (document.frm_pesquisa.campo8_ini_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo8_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo8_ini_m.value == '') || (document.frm_pesquisa.campo8_ini_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo8_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo8_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo8_ini_a.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo8_fim_d.value == '') || (document.frm_pesquisa.campo8_fim_d.value > 31)) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo8_fim_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo8_fim_m.value == '') || (document.frm_pesquisa.campo8_fim_m.value > 12)) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo8_fim_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo8_fim_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo8_fim_a.focus();
                            return false;
                        } else {
                            tam_dados += 1;
                        }
                    }
                } else {
                    if ((document.frm_pesquisa.campo8_ini_d.value != '') || (document.frm_pesquisa.campo8_ini_m.value != '') ||
						(document.frm_pesquisa.campo8_ini_a.value != '')) {
                        if ((document.frm_pesquisa.campo8_ini_d.value > 31) || (document.frm_pesquisa.campo8_ini_d.value == '')) {
                            alert('Digite um dia válido!');
                            document.frm_pesquisa.campo8_ini_d.focus();
                            return false;
                        } else if ((document.frm_pesquisa.campo8_ini_m.value > 12) || (document.frm_pesquisa.campo8_ini_m.value == '')) {
                            alert('Digite um mês válido!');
                            document.frm_pesquisa.campo8_ini_m.focus();
                            return false;
                        } else if (document.frm_pesquisa.campo8_ini_a.value == '') {
                            alert('Digite um ano válido!');
                            document.frm_pesquisa.campo8_ini_a.focus();
                            return false;
                        } else {
                            tam_dados += 1;
                        }
                    }
                }
            }
        }

        if (document.frm_pesquisa.campo1 != null) {
            if (document.frm_pesquisa.campo1.value.length > 0) {
                var sCampo1Pesquisa = document.frm_pesquisa.campo1.value;
                if ((sCampo1Pesquisa.substring(0, 1) == "'") && (sCampo1Pesquisa.substring(sCampo1Pesquisa.length - 1, sCampo1Pesquisa.length) == "'")) {
                    sCampo1Pesquisa = "\"" + sCampo1Pesquisa.substring(1, sCampo1Pesquisa.length - 1) + "\"";
                }
                document.frm_pesquisa.campo1.value = sCampo1Pesquisa;
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
        } else {     
            var sDadosPesquisa = document.frm_pesquisa.dados.value; 
            if ((sDadosPesquisa.substring(0, 1) == "'") && (sDadosPesquisa.substring(sDadosPesquisa.length - 1, sDadosPesquisa.length) == "'")) {
                sDadosPesquisa = "\"" + sDadosPesquisa.substring(1, sDadosPesquisa.length - 1) + "\"";
            }
            document.frm_pesquisa.dados.value = sDadosPesquisa;

            document.frm_pesquisa.submit.enabled = false;
            ajxFormPesquisaSubmit();
        }
    }
    return false;
}

function ValidaCamposDataVazio(iniDia, iniMes, iniAno, finDia, finMes, fimAno) {
    if (iniDia == '' && iniMes == '' && iniAno == '' && finDia == '' && finMes == '' && fimAno == '') {
        return false;
    } else {
        return true;
    }
}

function ContarAcessoMidia(codigoMidia) {
    $.ajax({
        type: "GET",
        url: 'asp/conta_acesso.asp?midia=' + codigoMidia,
        cache: false,
        success: function (data) {
        }
    });
}