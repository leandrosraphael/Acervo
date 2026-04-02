//###############################################################################//
// VARIAVEIS GLOBAIS                                                             //
//###############################################################################//
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

function LinkOrcamentos() {
	parent.mainFrame.location = "index."+ext+"?content=orcamento&iIdioma="+global_frame.iIdioma;
}

function LinkTrocaSenha() {
	parent.mainFrame.location = "index."+ext+"?content=troca_senha&iIdioma="+global_frame.iIdioma;
}

function LinkLogin() {
	var msg_login = getTermo(global_frame.iIdioma, 6228, 'Entrar', 0);
	abrePopup('login.'+ext+'?iIdioma='+global_frame.iIdioma,msg_login,320,220,false,true,"orcamento");
}

function LinkLogout() {
	parent.mainFrame.location = "logout."+ext+"?iIdioma="+global_frame.iIdioma+"&iBanner="+global_frame.iBanner;
}

function LinkCotacao(CodCotacao) {
	parent.mainFrame.location = "index."+ext+"?content=cotacao&cotacao="+CodCotacao+"&iIdioma="+global_frame.iIdioma;
}

function ValidaLogin() {
	var tam_codigo = document.login.codigo.value.length;
	if (tam_codigo == 0) {
		alert(getTermo(global_frame.iIdioma, 42, 'O login do usuário deve ser preenchido.', 0));
		document.login.codigo.focus();
		return false;
	}
	document.login.submit();
}

function FechaLogin() {
	parent.document.location='index.'+ext+'?content=orcamento&iBanner=' + global_frame.iBanner + '&iIdioma='+global_frame.iIdioma;
	parent.fechaPopup(); 
}

function validaTeclaLogin(campo, e) { 
	var key; 
	var tecla; 
	var validou = true;
	var strValidChars = "0813";
	var tecla = window.event ? e.keyCode : e.which; 
	if ( strValidChars.indexOf( tecla ) > -1 ) {
		validou = false;
	}
	key = String.fromCharCode( tecla); 
	if ( tecla == 13 ) {
		var tam_login = document.login.codigo.value.length;
		//var tam_senha = document.login.senha.value.length;
		if (tam_login == 0) {
			alert(getTermo(global_frame.iIdioma, 42, 'O login do usuário deve ser preenchido.', 0));
			document.login.codigo.focus();
			return false;
		}
		document.login.submit();		
	}
	return true;
}

function MascaraMoeda(objTextBox, SeparadorMilesimo, SeparadorDecimal, e){
    var sep = 0;
    var key = '';
    var i = j = 0;
    var len = len2 = 0;
    var strCheck = '0123456789';
    var aux = aux2 = '';
	if (e != null) {
		if(navigator.appName.indexOf("Netscape")!= -1) {
			var whichCode = e.which; 
		} else {
			var whichCode = e.keyCode; 
		}
		//Teclas Enter, Delete, Backspace
		if ((whichCode == 13) || (whichCode == 8) || (whichCode == 0))
			return true; 
		key = String.fromCharCode(whichCode); // Valor para o código da Chave
		if (strCheck.indexOf(key) == -1) return false; // Chave inválida
	}
    len = objTextBox.value.length;
    for(i = 0; i < len; i++)
        if ((objTextBox.value.charAt(i) != '0') && (objTextBox.value.charAt(i) != SeparadorDecimal)) break;
    aux = '';
    for(; i < len; i++)
        if (strCheck.indexOf(objTextBox.value.charAt(i))!=-1) aux += objTextBox.value.charAt(i);
    aux += key;
    len = aux.length;
    if (len == 0) objTextBox.value = '';
    if (len == 1) objTextBox.value = '0'+ SeparadorDecimal + '0' + aux;
    if (len == 2) objTextBox.value = '0'+ SeparadorDecimal + aux;
    if (len > 2) {
        aux2 = '';
        for (j = 0, i = len - 3; i >= 0; i--) {
            if (j == 3) {
                aux2 += SeparadorMilesimo;
                j = 0;
            }
            aux2 += aux.charAt(i);
            j++;
        }
        objTextBox.value = '';
        len2 = aux2.length;
        for (i = len2 - 1; i >= 0; i--)
        	objTextBox.value += aux2.charAt(i);
        objTextBox.value += SeparadorDecimal + aux.substr(len - 2, len);
    }
    return false;
}

function MascaraPorcentagem(objTextBox, SeparadorDecimal, e){
    var SeparadorMilesimo = '';
	var sep = 0;
    var key = '';
    var i = j = 0;
    var len = len2 = 0;
    var strCheck = '0123456789';
    var aux = aux2 = '';
	if (e != null) {
		if(navigator.appName.indexOf("Netscape")!= -1) {
			var whichCode = e.which; 
		} else {
			var whichCode = e.keyCode; 
		}
		//Teclas Enter, Delete, Backspace
		if ((whichCode == 13) || (whichCode == 8) || (whichCode == 0))
			return true; 
		key = String.fromCharCode(whichCode); // Valor para o código da Chave
		if (strCheck.indexOf(key) == -1) return false; // Chave inválida
	}
    len = objTextBox.value.length;
    for(i = 0; i < len; i++)
        if ((objTextBox.value.charAt(i) != '0') && (objTextBox.value.charAt(i) != SeparadorDecimal)) break;
    aux = '';
    for(; i < len; i++)
        if (strCheck.indexOf(objTextBox.value.charAt(i))!=-1) aux += objTextBox.value.charAt(i);
    aux += key;
    len = aux.length;
    if (len == 0) objTextBox.value = '';
    if (len == 1) objTextBox.value = '0'+ SeparadorDecimal + '0' + aux;
    if (len == 2) objTextBox.value = '0'+ SeparadorDecimal + aux;
    if ((len > 2) && (len < 5)) {
        aux2 = '';
        for (j = 0, i = len - 3; i >= 0; i--) {
            if (j == 3) {
                aux2 += SeparadorMilesimo;
                j = 0;
            }
            aux2 += aux.charAt(i);
            j++;
        }
        objTextBox.value = '';
        len2 = aux2.length;
        for (i = len2 - 1; i >= 0; i--)
        	objTextBox.value += aux2.charAt(i);
        objTextBox.value += SeparadorDecimal + aux.substr(len - 2, len);
    }
    return false;
}

function MascaraInt(objTextBox, e){
    var key = '';
    var strCheck = '0123456789';
	if (e != null) {
		if(navigator.appName.indexOf("Netscape")!= -1) {
			var whichCode = e.which; 
		} else {
			var whichCode = e.keyCode; 
		}
		//Teclas Enter, Delete, Backspace
		if ((whichCode == 13) || (whichCode == 8) || (whichCode == 0))
			return true; 
		key = String.fromCharCode(whichCode); // Valor para o código da Chave
		if (strCheck.indexOf(key) == -1) {
			return false; // Chave inválida
		} else {
			return true;
		}
	} else {
		return false;
	}
}

function atualizaCotacao(form) {
	var titulo = getTermo(global_frame.iIdioma, 5652, 'Cotação', 0);
	abrePopup('about:blank', titulo, 250, 120, false, false, 'orcamento');
	return true;
}

function enviaCotacao(acao, form) {
	if (acao === 'atualizar' || confirm(getTermo(global_frame.iIdioma, 6906, 'Ao finalizar o orçamento, este não poderá mais ser alterado. Deseja continuar?', 0))) {
		form.acao.value = acao;
		if (atualizaCotacao(form)) {
			form.submit();
		}
	}
}

function alteraIdioma(idioma) {
    var idiomaOld = global_frame.iIdioma;
    global_frame.iIdioma = idioma;
   
    var pos_idioma = main_frame.location.toString().indexOf("iIdioma=" + idiomaOld);
    if (pos_idioma >= 0) {
        main_frame.location = main_frame.location.toString().replace("iIdioma=" + idiomaOld, "iIdioma=" + idioma);
    } else {
        var pos = main_frame.location.toString().indexOf("?");
        if (pos < 0) {
            main_frame.location = main_frame.location + "?refresh_popup=1&iIdioma=" + idioma;
        } else {
            main_frame.location = main_frame.location + "&refresh_popup=1&iIdioma=" + idioma;
        }
    }
}

function helloPage() {
	$.ajax({
		type: 'GET',
		url: 'helloPage.' + ext + '?n=1',
	});
}