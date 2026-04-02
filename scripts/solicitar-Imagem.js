/* ANA - Solicitação de imagens */

function solicitarImagem() {
    var url, titulo;
    titulo = getTermo(global_frame.iIdioma, 9702, "Solicitação de imagens", 0);
    url = "asp/solicitar-imagem.asp";
    abrePopup(url, titulo, 800, 600, true, true);
}

function imagensConcedidas() {
    parent.mainFrame.location = "index.asp?veio_de=menu&content=imagem_concedida" + getGlobalUrlParams();
}

function formatarCPF() {
    $("#cpf-responsavel").attr("maxLength", "14");
    $("#cpf-responsavel").mask("999.999.999-99");
}

function formatarCNPJ() {
    $("#documento-instituicao").attr("maxLength", "18");
    $("#documento-instituicao").mask("99.999.999/9999-99");
}

function formatarCEP() {
    $("#cep-solicitante").attr("maxLength", "10");
    $("#cep-solicitante").mask("99.999-999");
}

function habilitarOutros() {
    if ($("#destinacaoTipoOutros").is(":checked"))
    {
        $("#label-destinacaooutros").text("Especificar outros*");
        $("#destinacao-outros").prop("disabled", false);
    }
    else
    {
        $("#label-destinacaooutros").text("Especificar outros");
        $("#destinacao-outros").prop("disabled", true);
    }
}

function alterarTipoDoc() {
    var Instituicao = $("#tipo-instituicao").val();
    if (Instituicao == 6)
    {
        $("#label-instituicao").text("Nome da instituição");
        formatarCPF();
        $("#nome-instituicao").val('');
        $("#nome-instituicao").prop("disabled", true);
        $("#dados-pj").css('display', 'none');
        $("#principal").css('width', '97%');        
    }
    else if (Instituicao < 6)
    {
        $("#label-instituicao").text("Nome da instituição*");
        formatarCNPJ();
        formatarCPF();
        $("#nome-instituicao").prop("disabled", false);
        $("#dados-pj").css('display', 'block');
        $("#principal").css('width', '95%');        
    }
}

function isEmpty(Str) {
    return ((Str == null) || (Str.length == 0))
}

function ValidarDadosSolicitacaoImagem(owner) {
    var aux;
    var Mensagem = "Os seguintes campos não foram preenchidos corretamente: \n";
    var retorno = true;

    aux = $("#tipo-instituicao").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo Tipo da instituição é obrigatório.";
        retorno = false;
    }

    if ($("#tipo-instituicao").val() != 6)
    {
        aux = $("#nome-instituicao").val();
        if (isEmpty(aux)) {
            Mensagem += "\nO campo Nome da instituição é obrigatório.";
            retorno = false;
        }

        aux = $("#documento-instituicao").val();

        if (isEmpty(aux)) {
            Mensagem += "\nO campo CNPJ é obrigatório.";
            retorno = false;
        } else {
            if (!isCNPJValid(aux)) {
                Mensagem += "\nO campo CNPJ é inválido.";
                retorno = false;
            }
        }
        aux = $("#endereco-solicitante").val();
        if (isEmpty(aux)) {
            Mensagem += "\nO campo Endereço do solicitante é obrigatório.";
            retorno = false;
        }

        aux = $("#cidade-solicitante").val();
        if (isEmpty(aux)) {
            Mensagem += "\nO campo Cidade do solicitante é obrigatório.";
            retorno = false;
        }

        aux = $("#uf-solicitante").val();
        if (isEmpty(aux)) {
            Mensagem += "\nO campo Estado do solicitante é obrigatório.";
            retorno = false;
        }

        aux = $("#cep-solicitante").val();
        if (isEmpty(aux)) {
            Mensagem += "\nO campo CEP do solicitante é obrigatório.";
            retorno = false;
        }
    }              
        
    aux = $("#cpf-responsavel").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo CPF é obrigatório.";
        retorno = false;
    } else {
        if (!isCPFValid(aux)) {
        	Mensagem += "\nO campo CPF é inválido.";
        	retorno = false;
        }
    }    

    aux = $("#nome-responsavel").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo Nome do responsável é obrigatório.";
        retorno = false;
    }

    aux = $("#email-responsavel").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo e-mail do responsável é obrigatório.";
        retorno = false;
    } else {
        var exreg = /^[a-zA-Z0-9][a-zA-Z0-9\._-]+@([a-zA-Z0-9\._-]+\.)[a-zA-Z-0-9]{2}/;
        // msmsms-99_.@msmsms-99_.br (se existir o final como .br tem q ter duas letras)
        if (! exreg.exec(aux)) {
            Mensagem += "\nO campo e-mail é inválido.";
            retorno = false;
        }
    }

    aux = $("#nome-cargo").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo Cargo do responsável é obrigatório.";
        retorno = false;
    }      

    aux = $("#rg-responsavel").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo RG do responsável é obrigatório.";
        retorno = false;
    }

    aux = $("#telefone-responsavel").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo Telefone do responsável é obrigatório.";
        retorno = false;
    }

    aux = $("#endereco-responsavel").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo Endereço do responsável é obrigatório.";
        retorno = false;
    }

    aux = $("#cidade-responsavel").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo Cidade do responsável é obrigatório.";
        retorno = false;
    }

    aux = $("#uf-responsavel").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo UF do responsável é obrigatório.";
        retorno = false;
    }

    aux = $("#cep-responsavel").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo CEP do responsável é obrigatório.";
        retorno = false;
    }

    aux = $("#finalidade-imagem").val();
    if (isEmpty(aux)) {
        Mensagem += "\nO campo Finalidade da imagem é obrigatório.";
        retorno = false;
    }
    
    aux = $("#input-destinacao > input:checked").length;
    if (aux <= 0) {
        Mensagem += "\nMarcar pelo menos um tipo de destinação da imagem.";
        retorno = false;
    }

    if ($("#destinacaoTipoOutros").is(":checked") && $("#destinacao-outros").val() == '')
    {
        Mensagem += "\nO campo Especificar outros da destinação da solicitação é obrigatório.";
        retorno = false;
    }

    if (! $("#politica-uso").is(":checked")) {
        Mensagem += "\nO campo sobre o aceite da Política de uso não foi marcado.";
        retorno = false;
    }

    if (! $("#politica-privacidade").is(":checked")) {
        Mensagem += "\nO campo sobre o aceite da Política de privacidade não foi marcado.";
        retorno = false;
    }

    if (retorno == false) {
        alert(Mensagem);
        return false;
    } else {
        parent.fechaPopup();
        owner.form.submit();
        return true;
    }
}

//Verifica se CPF é válido
function isCPFValid(strCPF) {
    var Soma;
    var Resto;
    Soma = 0;
    strCPF = strCPF.replace(/[^\d]+/g, '');

    if (strCPF == "00000000000") {
        return false;
    }

    for (i = 1; i <= 9; i++) {
        Soma = Soma + parseInt(strCPF.substring(i - 1, i)) * (11 - i);
    }

    Resto = (Soma * 10) % 11;
    if ((Resto == 10) || (Resto == 11)) {
        Resto = 0;
    }
        
    if (Resto != parseInt(strCPF.substring(9, 10))) {
        return false;
    }
        
    Soma = 0;
    for (i = 1; i <= 10; i++) {
        Soma = Soma + parseInt(strCPF.substring(i - 1, i)) * (12 - i);
    }

    Resto = (Soma * 10) % 11;
    if ((Resto == 10) || (Resto == 11)) {
        Resto = 0;
    }

    if (Resto != parseInt(strCPF.substring(10, 11))) {
        return false;
    }

    return true;
}

function isCNPJValid(cnpj) {

    cnpj = cnpj.replace(/[^\d]+/g, '');

    if (cnpj == '') {
        return false;
    }

    if (cnpj.length != 14) {
        return false;
    }
        

    // Elimina CNPJs invalidos conhecidos
    if (cnpj == "00000000000000" ||
        cnpj == "11111111111111" ||
        cnpj == "22222222222222" ||
        cnpj == "33333333333333" ||
        cnpj == "44444444444444" ||
        cnpj == "55555555555555" ||
        cnpj == "66666666666666" ||
        cnpj == "77777777777777" ||
        cnpj == "88888888888888" ||
        cnpj == "99999999999999") {
        return false;
    }
        

    // Valida DVs
    tamanho = cnpj.length - 2
    numeros = cnpj.substring(0, tamanho);
    digitos = cnpj.substring(tamanho);
    soma = 0;
    pos = tamanho - 7;
    for (i = tamanho; i >= 1; i--) {
        soma += numeros.charAt(tamanho - i) * pos--;
        if (pos < 2) {
            pos = 9;
        }
    }

    resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
    if (resultado != digitos.charAt(0)) {
        return false;
    }
        
    tamanho = tamanho + 1;
    numeros = cnpj.substring(0, tamanho);
    soma = 0;
    pos = tamanho - 7;
    for (i = tamanho; i >= 1; i--) {
        soma += numeros.charAt(tamanho - i) * pos--;
        if (pos < 2) {
            pos = 9;
        }
    }
    resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
    if (resultado != digitos.charAt(1)) {
        return false;
    }
        
    return true;
}

function LinkPaginacaoImagemAno(IrPara) {
    var Filtro;
    var anoAtual;
    var Frame;
    var url;

    if (parent.hiddenFrame != null) {
        Frame = parent.hiddenFrame;
    } else if (parent.parent.hiddenFrame != null) {
        Frame = parent.parent.hiddenFrame;
    } else {
        Frame = parent.parent.parent.hiddenFrame;
    };
    if (IrPara == "primeiro") {
        Frame.idx_vetor_ano = 0;
        Filtro = Frame.img_vetor_ano[0];
    } else if (IrPara == "ultimo") {
        Frame.idx_vetor_ano = Frame.img_vetor_ano.length-1;
        Filtro = Frame.img_vetor_ano[Frame.idx_vetor_ano];
    } else if (IrPara == 'proximo') {
        var idx;
        idx = Frame.idx_vetor_ano;
        idx--;
        if (idx < 0) {
            AtualizarPaginacao();
            return false;
        }
        Filtro = Frame.img_vetor_ano[idx];
        Frame.idx_vetor_ano = idx;
    } else if (IrPara == 'anterior') {
        var idx;
        idx = Frame.idx_vetor_ano;
        idx++;
        if (idx > (Frame.img_vetor_ano.length - 1)) {
            AtualizarPaginacao();
            return false;
        }
        Filtro = Frame.img_vetor_ano[idx];
        Frame.idx_vetor_ano = idx;
    } else {
        Filtro = IrPara;
        var i;
        for (i = 0; i < Frame.img_vetor_ano.length; i++) {
            if (Frame.img_vetor_ano[i] == Filtro) {
                Frame.idx_vetor_ano = i;
                break;
            }
        }
    }

    document.body.style.cursor = "wait";

    Frame.img_pagina_atual = 1;

    ListarFichaImagens(Filtro, 1);
    AtualizarPaginacao();

    document.body.style.cursor = "default";
}

function ListarFichaImagens(Ano, Pagina) {
    $.ajax({
        type: 'POST',
        url: ext + '/ajxFichasImagens.asp?Ano=' + Ano + '&Pagina=' + Pagina + getGlobalUrlParams(),
        data: "",
        success: function (data) {
            $('#div-ficha-imagem').html(data);
        }
    });
}

function AtualizarPaginacao() {
    var Frame;

    if (parent.hiddenFrame != null) {
        Frame = parent.hiddenFrame;
    } else if (parent.parent.hiddenFrame != null) {
        Frame = parent.parent.hiddenFrame;
    } else {
        Frame = parent.parent.parent.hiddenFrame;
    };

    var numAnos = Frame.img_vetor_ano.length;
    var IndiceAno = Frame.idx_vetor_ano;
    var TotalRegistros = Frame.img_total_registros_ano;
    var TotalImagens = Frame.img_total_imagens;
    var paginaAtual = Frame.img_pagina_atual;
    var NumPaginas = Frame.img_paginas;
    var anoAtual = Frame.img_vetor_ano[Frame.idx_vetor_ano];
    var Anos = "";
    
    for (var i = 0; i < Frame.img_vetor_ano.length; i++) {
        if (Anos != "") {
            Anos += ',';
        }
        Anos += Frame.img_vetor_ano[i];
    }

    var urlPag = ext + '/ajxAtualizaPaginacao.asp?AnoAtual='+ anoAtual + '&listaAnos=' + Anos + '&numAnos=' + numAnos + '&indiceAno=' + IndiceAno + '&numPaginas=' + NumPaginas + '&paginaAtual=' + paginaAtual + '&totalRegistros=' + TotalRegistros + '&totalImagens=' + TotalImagens;
    $.ajax({
        type: 'POST',
        url: urlPag,
        data: "",
        success: function (data) {
            $('#paginacao-imagem-topo').html(data);
            $('#paginacao-imagem-rodape').html(data);
        }
    });
}

function LinkPaginacaoImagemReg(IrPara) {
    var Filtro;
    var anoAtual;
    var Frame;

    if (parent.hiddenFrame != null) {
        Frame = parent.hiddenFrame;
    } else if (parent.parent.hiddenFrame != null) {
        Frame = parent.parent.hiddenFrame;
    } else {
        Frame = parent.parent.parent.hiddenFrame;
    };

    if (IrPara == 'anterior') {
        var idx;
        idx = Frame.img_pagina_atual;
        idx--;
        if (idx < 1) {
            return false;
        }
        Frame.img_pagina_atual = idx;
    } else if (IrPara == 'proximo') {
        var idx;
        idx = Frame.img_pagina_atual;
        idx++
        if (idx > Frame.img_paginas) {
            return false;
        }
        Frame.img_pagina_atual = idx;
    } else {
        Frame.img_pagina_atual = IrPara;
    }

    anoAtual = Frame.img_vetor_ano[Frame.idx_vetor_ano];

    document.body.style.cursor = "wait";

    ListarFichaImagens(anoAtual, Frame.img_pagina_atual);
    AtualizarPaginacao();
    document.body.style.cursor = "default";
}