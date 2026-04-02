<!DOCTYPE HTML>
<html>
<head>	
	<!-- #include file="../config.asp" -->

	<title></title>

	<script src="../scripts/jquery-1.7.1.min.js" type="text/javascript"></script>
	<script src="../scripts/zoom.js" type="text/javascript"></script>

	<% if (config_customizacao_layout = 0) then %>
		<link href="../inc/estilo.css" rel="stylesheet" type="text/css">
	<% else %>
		<link href="../libasp/estilo.asp" rel="stylesheet" type="text/css">
	<% end if %>	
	
	<link href="../inc/zoom.css" rel="stylesheet" type="text/css">
</head>
<body>
	<%
		codItem = Request.QueryString("item")
		codImagem = Request.QueryString("imagem")
		zoom = Request.QueryString("zoom")
		content = Request.QueryString("content")
		
		if (zoom = "") then
			zoom = "1"
		end if
		
		if(zoom = "1") then
	%>	
	<div id="divNav">
		<input type="button" id="maisButton" class="botao-zoom" value="[+]" />
		<input type="button" id="menosButton" class="botao-zoom" value="[-]" />
		<input type="button" id="originalButton" value="TAMANHO ORIGINAL" />
		<input type="button" id="ajusteButton" value="AJUSTAR" />
	</div>	
	<div id="divImg">
	<%
		end if

		sUrl = "imagem.asp?item=" & codItem + "&imagem=" & codImagem & "&zoom=" & zoom & "&content=" & content
		Response.Write "<img id='img' src='" & sUrl & "' border='0'>"	
	%>
	</div>	
</body>
</html>