<%
	'Desabilitando cache pois estava influenciando na passagem de parâmetros no site
	'Ao recarregar a mesma página os dados nem sempre eram atualizados
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<!DOCTYPE HTML>
<!-- #include file="libasp/funcoes.asp" -->
<!-- #include file="config.asp" -->
<!-- #include file="libasp/roclient.asp" -->
<!-- #include file="libasp/header.asp" -->
<!-- #include file="inc/skin.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
	<% if (config_customizacao_layout = 0) then %>
		<link href="inc/estilo.css" rel="stylesheet" type="text/css">
	<% else %>
		<link href="libasp/estilo.asp" rel="stylesheet" type="text/css">
	<% end if %>

	<script type="text/javascript">
		var buscaUnificada = '<%= Request.QueryString("buscaunificada") %>';
		var content = '<%= Request.QueryString("content") %>';
		var dados = '<%= Request.QueryString("dados") %>';
		var tmp = '<%= Request.QueryString("tmp") %>';
		var tmp_objeto = '<%= Request.QueryString("tmp_objeto") %>';

		temParametro = buscaUnificada != null;
		if (temParametro && parent.hiddenFrame == undefined) {
			window.sessionStorage.setItem("content", content);
			window.sessionStorage.setItem("dados", dados);
			window.sessionStorage.setItem("tmp", tmp);
			window.sessionStorage.setItem("tmp_objeto", tmp_objeto);

			window.location = 'index.html';
		}
	</script>

	<script type="text/javaScript" src="scripts/jquery-1.7.1.min.js"></script>
	<script type="text/javaScript" src="libasp/funcoes.js?b=<%=build_web%>"></script>

</head>

<body leftmargin="0" topmargin="0">

<div id='dvLoad' style="position:absolute; width: 100%; top: 150px; display: none">
    <center>
        <table width="130px" class="grid">
            <tr><td class="td_valor">
                <center><br><img src="imagens/mozilla_blu.gif" border="0"><br><br>Carregando<br><br></center>
            </td></tr>
        </table>
    </center>
</div>
<div id='screen'></div>
<div id='DivLogin' class='DivLogin'></div>


<div id="divNewMain">
	<table width="970" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="td_cabecalho_banner esquerda" style="vertical-align:middle; padding-left: 42px;">
					
			<span>&nbsp;<%=config_nome_cliente%>&nbsp;</span>
			
		</td>

	</tr>
	<tr> 
	<td class="td_corpo_externo">
	<br>
	<table class="td_corpo" border="0" cellpadding="1" cellspacing="0">
	<tr>
	<td valign="top" align="center">
	<div id="ajxDiv">
	<%

	numero_serie = 0

	if (sMsgErro <> "") Then
		Response.Write sMsgErro
	else
		numero_serie = ROService.GetNumeroDTA
		bAutenticado = true
        'VALIDA SE UTILIZA O SINGLE SIGN ON
		if (numero_serie = 6209) then
			if (Session("logado") = "true") then
				Session("Autenticado") = "1"
			else
				bAutenticado = false
			end if
		else
			Session("UsaLogin") = ROService.UsaLogin

			if (Session("UsaLogin") = 1) then
				if Session("Autenticado") <> "1" then		
					chave = Request.Form("chave")
					if ROService.ValidaLogin(chave) = false then
						bAutenticado = false
					else
						Session("Autenticado") = "1"
					end if
    			end if
			end if
		end if

		if bAutenticado then
			Select case (Request.Querystring("content"))
				case "pesquisa"
					%>
					<!-- #include file='asp/pesquisa.asp' -->
					<%
				case "detalhe_resultados"
					%>
					<!-- #include file='asp/resumo.asp' -->
					<%
				case "resultados"
					%>
					<!-- #include file='asp/resultado.asp' -->
					<%
				case "detalhe"
					%>
					<!-- #include file='asp/detalhes.asp' -->
					<%
				case else
					%>
					<!-- #include file='asp/pesquisa.asp' -->
					<%
			End Select
		else
			if (numero_serie = 6209) and (mensagem <> "") then
				Response.Write "<img src='imagens/erro.gif'>&nbsp;<b>"&mensagem&"</b>" 
			else
				Response.Write "<img src='imagens/erro.gif'>&nbsp;<b>Sessão não autenticada.</b>"
			end if
		end if
	End if
	%>
	</div>
	</td>
	</tr>
	</table>
	<br>
	</td>
	</tr>
	</table>
<center>
<% if (numero_serie = 4795) then %>
	<br>
	<a class="link_prima" href="http://www4.tjrj.jus.br/museu/index.html" target="_blank"><b><u>Acervo Bibliográfico</b></u></a>
<%end if%>
<br>
<font color="#FFFFFF">
Desenvolvido por</font>
<a class="link_prima" href="http://www.prima.com.br" target="_blank">Prima</a> 
<br><br>
</center>
<script type="text/javascript">
    var mainResize = function () {
        var winH = $(window).height();
        var mainH = $('#divNewMain').height();

        if (winH > mainH) {
            $('#divNewMain').css('min-height', winH + 'px');
        }
    };

    $(function () {
        $(window).resize(mainResize);
        mainResize();
    });

</script>
</body>
</html>
<%
Set ROService = nothing
Set ROServer = nothing
%>

