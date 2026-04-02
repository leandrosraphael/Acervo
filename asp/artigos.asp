<% Response.Buffer=True %>
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
<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" /> 
<link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />
<% if config_css_estatico = 1 then %>
	<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
<% else %>
	<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
<% end if %>
    <link href="../inc/contraste.css" rel="stylesheet" type="text/css" />   
    <link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
<%
    iCodEx    = Request.QueryString("codex") 
	iIndexSrv = IntQueryString("Servidor", 1)
	iCodObra  = Request.QueryString("obra") 

if iCodObra = "" then %>
	<title><%=getTermo(global_idioma, 1339, "Artigos", 0)%></title>
<% else %>
	<title><%=getTermo(global_idioma, 53, "Analíticas", 0)%></title>
<% end if
	if Request.QueryString("popup") = "1" then
		Veio_Popup = true
	else 
		Veio_Popup = false
	end if
		
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 	
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
%>
</head>
<body class="body_artigo">
<p class="centro">
<div style="padding: 6px; width: 953px; height: 429px; position: relative">
	<%

	if iCodObra = "" then 
	
		'Busca os dados do título do artigo
		Set ROService = ROServer.CreateService("Web_Consulta")		
		Set objDados  = ROServer.CreateComplexType("TDadosPerArtigo")
		Set objDados  = ROService.DadosPerArtigo(iCodEx)

		'Trata resultado
		sTit_Per	 = objDados.sTitulo
		sAno		 = objDados.sAno
		sVol		 = objDados.sVolume
		sNum		 = objDados.sNumero
		sParte  	 = objDados.sParte
		iPrim_Artigo = objDados.iArtigo
		sPrim_Titulo = objDados.sArtigo
	
		Set objDados = nothing			
		Set ROService = nothing
		TrataErros(1)

		Response.write "<table class='max_width'>"
		Response.write "<tr><td style='text-align: left' class='popup_cabecalho_background'>&nbsp;<b><span style='color: #FFFFFF'>"&getTermo(global_idioma, 465, "Periódico", 0)&"</span></b>&nbsp;</td></tr>"
		Response.write "<tr><td style='text-align: left' class='td_detalhe_valor'>&nbsp;"&sTit_Per&"<br/>"	
	
		'Monta dados do exemplar
		sDados = ""	
		'Ano
		if (sAno <> "") then 
			sDados = sDados & "<b>"&getTermo(global_idioma, 67, "Ano",0)&":</b> "&sAno&vbcrlf 
		end if
		'Volume
		if (sVol <> "") then  	
			if sDados <> "" then
				sDados = sDados & "&nbsp;&nbsp;"
			end if
			sDados = sDados & "&nbsp;&nbsp;<b>"&getTermo(global_idioma, 66, "Volume", 0)&":</b> "&sVol&vbcrlf	
		end if
		'Numero
		if (sNum <> "") then
			if sDados <> "" then
				sDados = sDados & "&nbsp;&nbsp;"
			end if
			sDados = sDados & "&nbsp;&nbsp;<b>"&getTermo(global_idioma, 101, "Número", 0)&":</b> "&sNum&vbcrlf
		end if
		'Parte
		if (sParte <> "") then
			if sDados <> "" then
				sDados = sDados & "&nbsp;&nbsp;"
			end if
			sDados = sDados & "&nbsp;&nbsp;<b>"&getTermo(global_idioma, 1340, "Parte", 0)&":</b> "&sParte&vbcrlf
		end if

		Response.write "&nbsp;"&sDados&"</td></tr>"	
		Response.Write "</table><br />" &vbcrlf
		if Veio_Popup = true then
			Response.Write "<iframe name='artigos' style='border: 0' scrolling='auto' width='27%' height='80%' src='lista_artigos.asp?popup=1&codex="&iCodEx&"&Servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'></iframe>"
			Response.Write "<iframe name='detalhe' style='border: 0' scrolling='auto' width='73%' height='80%' src='detalhe_artigo.asp?popup=1&codex="&iCodEx&"&Codigo="&iPrim_Artigo&"&Servidor="&iIndexSrv&"&tit_artigo="&sPrim_Titulo&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'></iframe>"
		else
			Response.Write "<iframe name='artigos' style='border: 0' scrolling='auto' width='27%' height='80%' src='lista_artigos.asp?codex="&iCodEx&"&Servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'></iframe>"
			Response.Write "<iframe name='detalhe' style='border: 0' scrolling='auto' width='73%' height='80%' src='detalhe_artigo.asp?codex="&iCodEx&"&Codigo="&iPrim_Artigo&"&Servidor="&iIndexSrv&"&tit_artigo="&sPrim_Titulo&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'></iframe>"
		end if
	else
	
		'Busca os dados do título da analitica
		Set ROService = ROServer.CreateService("Web_Consulta")		
		Set objDados  = ROServer.CreateComplexType("TDadosObraAna")
		Set objDados  = ROService.DadosObraAna(iCodObra)

		'Trata resultado
		sTit_Obra	 = objDados.sTitulo
		iPrim_Ana    = objDados.iAnalitica
		sPrim_Titulo = objDados.sAnalitica
	
		Set objDados = nothing			
		Set ROService = nothing
		
		TrataErros(1)

		Response.write "<table class='max_width'>"
		Response.write "<tr><td style='text-align: left' class='popup_cabecalho_background'>&nbsp;<b><span style='color: #FFFFFF'>"&getTermo(global_idioma, 464, "Obra", 0)&"</span></b>&nbsp;</td></tr>"
		Response.write "<tr><td class='td_detalhe_valor esquerda'>&nbsp;"&sTit_Obra&"<br />&nbsp;"	
	
		Response.write "</td></tr>"	
		Response.Write "</table><br />" &vbcrlf

		Response.Write "<iframe name='artigos' style='border: 0' scrolling='auto' width='27%' height='80%' src='lista_artigos.asp?obra="&iCodObra&"&Servidor="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'></iframe>"
		Response.Write "<iframe name='detalhe' style='border: 0' scrolling='auto' width='73%' height='80%' src='detalhe_artigo.asp?obra="&iCodObra&"&Codigo="&iPrim_Ana&"&Servidor="&iIndexSrv&"&tit_artigo="&sPrim_Titulo&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"'></iframe>"
	end if
%>	
</div>
</p>
</body>
</html>