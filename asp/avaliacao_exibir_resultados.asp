<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
    <head>
		
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=edge" />
		<% sDiretorioArq = "asp" %>
		
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
		
        <link href="../inc/avaliacao_exibir_resultados.css" rel="stylesheet" type="text/css" /> 
        <link href="../inc/contraste.css" rel="stylesheet" type="text/css" /> 
        <link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
		
        <%
		
		iCodigo  = request.QueryString("Codigo")

		'arMySel = split(Session(sSessao),",")
	
	    On Error Resume Next
	    SET ROService = ROServer.CreateService("Web_Consulta")
	
	    xml_AvaliacaoOnline = ROService.DetalhesAvaliacaoOnLine(iCodigo)
	
	    TrataErros(1)
	
	    SET ROService = nothing
	
	    if left(xml_AvaliacaoOnline,5) = "<?xml" then
		    Set xmlDoc = CreateObject("Microsoft.xmldom")
		    xmlDoc.async = False
		    xmlDoc.loadxml xml_AvaliacaoOnline
		    Set xmlRoot = xmlDoc.documentElement
		    '************************************************************
		    ' INFORMAÇÂO DO TITULO
		    '************************************************************
		    if xmlRoot.nodeName = "avaliacoes" then

				avaliacaoMedia = xmlRoot.attributes.getNamedItem("media").value
				avaliacaoTotal = xmlRoot.attributes.getNamedItem("quantidade").value

				dim nota(4), quantidade(4), contador
				contador = 0

			    For Each xmlNota In xmlRoot.childNodes

					'************************************************************
				    ' AVALIAÇÃO
 				    '************************************************************
				    if xmlNota.nodeName = "avaliacao" then
					    nota(contador) = xmlNota.attributes.getNamedItem("nota").value
						quantidade(contador) = xmlNota.attributes.getNamedItem("quantidade").value
						contador = contador + 1
				    end if
			    Next
		    end if
	    end if

		
	    On Error Resume Next
	    SET ROService = ROServer.CreateService("Web_Consulta")
	
	    xml_titulo = ROService.RetornaTitulo(0, iCodigo)
	
	    TrataErros(1)
	
	    SET ROService = nothing
	
	    if left(xml_titulo,5) = "<?xml" then
		    Set xmlDoc = CreateObject("Microsoft.xmldom")
		    xmlDoc.async = False
		    xmlDoc.loadxml xml_titulo
		    Set xmlRoot = xmlDoc.documentElement
		    '************************************************************
		    ' INFORMAÇÂO DO TITULO
		    '************************************************************
		    if xmlRoot.nodeName = "INFO_TITULO" then
			    For Each xmlPNome In xmlRoot.childNodes
				    '************************************************************
				    ' TITULO
 				    '************************************************************
				    if xmlPNome.nodeName = "TITULO" then
					    titulo = xmlPNome.attributes.getNamedItem("VALOR").value
				    end if
			    Next
		    end if
	    end if
		 
		 %>
		<title><% Response.Write getTermo(global_idioma, 6693, "Avaliação", 0) & " - " & getTermo(global_idioma, 6699, "Votar", 0) %></title>

	</head>

     <body class="popup" onload="parent.fechaLoadingPopup();">
	    
		<div class="divPrincipal">
		
			<b><%=titulo%></b>

				<div class="divPrincipal">

					<div class="divConteudo div-media">
						<div class="divAvaliacaoMedia">
							<b><% =getTermo(global_idioma, 6695, "Avaliação média", 0) %></b><br />
							<span class="mediaAvaliacao"><%=avaliacaoMedia%></span>
							<p><% =ImprimeEstrelas(avaliacaoMedia,true) %></p>
							<p><% if avaliacaoTotal = 1 then
								  Response.Write avaliacaoTotal & " " & getTermo(global_idioma, 6693, " avaliação", 0)
								  else
								  Response.Write avaliacaoTotal & " " & getTermo(global_idioma, 6694, " avaliações", 0)
								  end  if
								%>
							</p>
						</div>
					</div>

					<div class="divConteudo">
						<%
							Dim i
							For i = 1 to 5
							Response.Write "<p><span class='opcaoVoto'>" & ImprimeEstrelas(i & ",0" ,true) & "</span></p>"
							next
						%>
					</div>
					
					<div class="divConteudo div-quantidades">
						<%
							Dim f
							For f = 0 to 4
							Response.Write "<p><span class='opcaoVoto'>(" & quantidade(f) & ")</span></p>"
							next
						%>
					</div>

				</div>

		</div>

		 <div class="div-BotaoFechar">
			 <input class="button_termo" type="button" onclick="javascript:parent.fechaPopup('asp');" value="<%=getTermo(global_idioma, 220, "Fechar", 0)%>" id="fechar"/>
		</div>

    </body>

</html>
