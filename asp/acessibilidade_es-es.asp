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
	<b>Accesibilidad en SophiA Biblioteca</b>
</p>
<div style="text-align: justify;margin:10px;">
	<br />
	<br />
	La accesibilidad es una característica esencial que conlleva una mejora en la calidad de vida de las personas. Por medio de ella es posible que personas con 
    discapacidades o limitaciones físicas puedan acceder a servicios, productos e informaciones en sistemas informáticos y de comunicación. Los patrones de accesibilidad 
    utilizados en SophiA Biblioteca y en su OPAC Web de consulta siguen las principales recomendaciones de la W3C (World Wide Web Consortium).
    <br />
    <br />
	<b>Contraste</b>
	<br />
    <br />
    En la parte superior derecha del OPAC Web encontrará la  opción “Alto contraste”.  Esta opción permite una lectura confortable a usuarios
    con baja visión, daltonismo o personas que utilizan monitores monocromáticos. Basta con clicar en el link para cambiar el contraste del OPAC Web. 
    Para volver a la visualización normal, basta con pulsar nuevamente en el link y la apariencia original será restablecida.
	<br />
    <br />
	<b>Alteración del tamaño de las fuentes</b>
	<br />
    <br />
    Los navegadores permiten que las fuentes puedan ampliarse o disminuirse. Para realizar estas acciones, utilice las siguientes teclas:
    <br />
    <br />

    <div style="margin-left:10px;">
	    <table style="width:100%" class="td_acessibilidade">
            <tr>
                <th><b>Acción</b></th>
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
    Es posible pulsar las teclas repetidas veces hasta alcanzar el tamaño deseado. Esta funcionalidad es válida en los navegadores 
    Chrome, Internet Explores, Firefox, Ópera y Safari.

</div>
</body>
</html>