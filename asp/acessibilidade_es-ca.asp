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
	<b>Accessibilitat en SophiA Biblioteca</b>
</p>
<div style="text-align: justify;margin:10px;">
	<br />
	<br />
	L'accessibilitat és una característica essencial que comporta una millora en la qualitat de vida de les persones. Per mitjà d'ella és possible que 
    persones amb discapacitats o limitacions físiques puguin accedir a serveis, productes i informacions en sistemes informàtics i de comunicació. 
    Els patrons d'accessibilitat utilitzats en SophiA Biblioteca i en la seva OPAC Web de consulta segueixen les principals recomanacions de la W3C (World Wide Web Consortium).
    <br />
    <br />
	<b>Contrast</b>
	<br />
    <br />
    En la part superior dreta del OPAC Web trobarà l'opció “Alt contrast”. Aquesta opció permet una lectura confortable a usuaris amb baixa visió, 
    daltonisme o persones que utilitzen monitors monocromàtics. N'hi ha prou amb clicar en el link per canviar el contrast del OPAC Web. 
    Per tornar a la visualització normal, n'hi ha prou amb prémer novament en el link i l'aparença original serà restablida.
	<br />
    <br />
	<b>Alteració de la grandària de les fonts</b>
	<br />
    <br />
    Els navegadors permeten que les fonts puguin ampliar-se o disminuir-se. Per realitzar aquestes accions, utilitzi les següents tecles:
    <br />
    <br />

    <div style="margin-left:10px;">
	    <table style="width:100%" class="td_acessibilidade">
            <tr>
                <th><b>Acció</b></th>
                <th><b>Windows</b></th>		
                <th><b>Mac</b></th>
            </tr>
            <tr>
                <td>Ampliar pantalla</td>
                <td><b>CTRL</b><b style="font-size:18px;"> + </b></td>	
                <td><b>COMMAND</b><b style="font-size:18px;"> + </b></td>
            </tr>
            <tr>
                <td>Disminuir pantalla</td>
                <td><b>CTRL</b><b style="font-size:18px;"> - </b></td>
                <td><b>COMMAND</b><b style="font-size:18px;"> - </b></td>
            </tr>
        </table>	
	</div>

    <br />
    <br />
    És possible prémer les tecles repetides vegades fins a aconseguir la grandària desitjada. Aquesta funcionalitat és vàlida en els navegadors 
    Chrome, Internet Exploris, Firefox, Òpera i Safari.

</div>
</body>
</html>