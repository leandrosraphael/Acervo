function BuscarDadosAceleradorBusca(grupo) {
    $.ajax({
        type: 'GET',
        url: 'asp/ajxBuscarDadosAceleradorBusca.asp?grupo=' + grupo,
        success: function (data) {
            $('#painel_' + grupo).html(data);
        }
    });
}

function LimparBuscaCombinada() {
    Resetar();
    atribuiComb();
    move_layer('combinada', parent.hiddenFrame.layerX);
}

function RetornarResultadoDeBuscaDoAcelerador(codigo, grupo) {
    LimparBuscaCombinada();

    global_frame.tipoAcelerador = grupo;

    if (global_frame.tipoAcelerador === "material") {
        global_frame.comb_material = codigo;
    } else {
        global_frame.comb_reposdigital = codigo;
    }

    var url = "index.asp?acelerador=" + grupo + "&modo_busca=combinada&content=resultado&submeteu=combinada&rapida_campo=palavra_chave";
    parent.mainFrame.location = url + getGlobalUrlParams();
}

$(function () {
    BuscarDadosAceleradorBusca("material");
    BuscarDadosAceleradorBusca("repositorio");
});