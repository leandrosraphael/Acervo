<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
<%
sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />

<% if config_css_estatico = 1 then %>
	<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
<% else %>
	<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
<% end if %>

<link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
<script type='text/javascript' src='../scripts/ajxSeries.js'></script>
<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>

</head>
<body class="popup" style="padding-top: 6px;">
<script type="text/javascript">parent.fechaLoadingPopup();</script>
<br />
<br />
<p class="centro">
	<b>Acessibilidade no SophiA Biblioteca</b>
</p>
<div style="text-align: justify;margin:10px;">
	<br />
	<br />
	A acessibilidade é característica essencial que garante a melhoria da qualidade de vida das pessoas. Por meio dela é possível a pessoas com deficiências	
	ou limitações físicas a participação em atividades, serviços, produtos e informações, inclusive nos sistemas de tecnologia e comunicação. Os padrões de  
	acessibilidade utilizados no SophiA Biblioteca em seu Terminal de consulta são as principais recomendações do W3C (World Wide Web Consortium).
    <br />
    <br />
	<b>Contraste</b>
	<br />
    <br />
    Na parte superior do Terminal Web está presente a opção de alteração do contraste da tela. Essa alteração permite leitura confortável a usuários com baixa
    visão, daltonismo ou pessoas que utilizam monitores monocromáticos. Basta clicar no link para alterar o contraste do Terminal Web, eliminando as informações de cor. 
    Para retornar à visualização normal, basta clicar novamente no link que a aparência original será restabelecida.
	<br />
    <br />
	<b>Alteração do tamanho das fontes</b>
	<br />
    <br />
    Os navegadores permitem que as fontes sejam ampliadas ou diminuídas. Para realizar essas ações utilize as seguintes teclas:
    <br />
    <br />

    <div style="margin-left:10px;">
	    <table style="width:100%" class="td_acessibilidade">
            <tr>
                <th><b>Ação</b></th>
                <th><b>Windows</b></th>		
                <th><b>Mac</b></th>
            </tr>
            <tr>
                <td>Ampliar tela</td>
                <td><b>CTRL</b><b style="font-size:18px;"> + </b></td>	
                <td><b>COMMAND</b><b style="font-size:18px;"> + </b></td>
            </tr>
            <tr>
                <td>Diminuir tela</td>
                <td><b>CTRL</b><b style="font-size:18px;"> - </b></td>	
                <td><b>COMMAND</b><b style="font-size:18px;"> - </b></td>
            </tr>
        </table>	
	</div>

    <br />
    <br />
    É possível pressionar as teclas repetidas vezes, até alcançar o tamanho desejado. Essa funcionalidade é utilizada para os navegadores Chrome, Internet Explorer, Firefox, Ópera e Safari.

</div>
</body>
</html>