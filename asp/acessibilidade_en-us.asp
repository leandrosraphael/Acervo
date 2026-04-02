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
	<b>Accessibility in SophiA Library</b>
</p>
<div style="text-align: justify;margin:10px;">
	<br />
	<br />
	Accessibility is a key feature that ensures the improvement of people's quality of life. Through it is possible for people with disabilities or 
    physical limitations to participate in activities, services, products and information, including technology and communication systems. 
    Accessibility standards used in SophiA Library in the OPAC are the main W3C (World Wide Web Consortium) recommendations.
    <br />
    <br />
	<b>Contrast</b>
	<br />
    <br />
    At the top of OPAC is available the option to display the screen contrast. This change allows comfortable reading for users with impaired vision, 
    color blindness or people using monochrome monitors. Just click the link to change the OPAC contrast, eliminating the color information. 
    To return to normal view, just click again on the link and the original appearance is restored.
	<br />
    <br />
	<b>Changing the size of fonts</b>
	<br />
    <br />
    Browsers allow fonts to be enlarged or reduced. Simply use the following keys:
    <br />
    <br />

    <div style="margin-left:10px;">
	    <table style="width:100%" class="td_acessibilidade">
            <tr>
                <th><b>Action</b></th>
                <th><b>Windows</b></th>		
                <th><b>Mac</b></th>
            </tr>
            <tr>
                <td>Larger screen</td>
                <td><b>CTRL</b><b style="font-size:18px;"> + </b></td>		
                <td><b>COMMAND</b><b style="font-size:18px;"> + </b></td>
            </tr>
            <tr>
                <td>Decrease screen</td>
                <td><b>CTRL</b><b style="font-size:18px;"> - </b></td>	
                <td><b>COMMAND</b><b style="font-size:18px;"> - </b></td>
            </tr>
        </table>	
	</div>

    <br />
    <br />
    You can press the keys repeated times, until you reach the desired size. This functionality is used for the browsers: Chrome, Internet Explorer, Firefox, Opera and Safari.

</div>
</body>
</html>