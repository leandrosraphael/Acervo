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
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title></title>
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
<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
</head>
<body class="body_artigo">
<script type="text/javascript">parent.parent.fechaLoadingPopup();</script>
<%
	if Request.QueryString("popup") = "1" then
		Veio_Popup = true
	else 
		Veio_Popup = false
	end if

	iArtigo   	= IntQueryString("Codigo", 0)
	sTit_Artigo = Server.HTMLEncode(Request.QueryString("tit_artigo"))
	iCodObra   	= IntQueryString("obra", 0)
	iCodEx   	= IntQueryString("codex", 0)
	iIndexSrv = IntQueryString("Servidor", 1)
		
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 	
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	
	if (iArtigo <> "") then
		conta_informacoes = 0 'Se não encontar nenhuma informação exibir texto "Nenhuma Informação encontrada"
		Response.write "<table class='max_width'>"
		Response.write "<tr><td colspan='2' class='td_detalhe_descricao' style='text-align: center; width: auto'>&nbsp;<b>"&sTit_Artigo&"</b>&nbsp;</td></tr>"
		
	  	'Busca dados do artigo
		Set ROService = ROServer.CreateService("Web_Consulta")	
		sXML = ROService.MontaArtigo(iArtigo,global_idioma)
		Set ROService = nothing
		TrataErros(1)
	
		'Processa XML com títulos dos artigos
		if (left(sXML,5) = "<?xml") then
			
			'Processa o XML com os títulos dos artigos		
			Set xmlDoc = CreateObject("Microsoft.xmldom")
			xmlDoc.async = False
			xmlDoc.loadxml sXML
			Set xmlRoot = xmlDoc.documentElement
			
			For Each xmlArtigo In xmlRoot.childNodes
				'Processa assuntos
				if xmlArtigo.nodeName  = "ASSUNTOS" then
					sAux = ""
					For Each xmlAssuntos In xmlArtigo.childNodes				
						if xmlAssuntos.nodeName  = "ASSUNTO" then
							sAssunto = xmlAssuntos.attributes.getNamedItem("Valor").value					
							sSeq     = xmlAssuntos.attributes.getNamedItem("Seq").value
							
							if sAux <> "" then
								sAux = sAux & "<br/>"
							end if
									
							sAux = sAux & sSeq & "&nbsp;" & sAssunto
						end if
					Next
					
					Response.Write "<tr>"
					Response.Write "<td class='td_detalhe_descricao' style='width: 90px'>&nbsp;"&getTermo(global_idioma, 205, "Assuntos", 0)&"&nbsp;</td>"
					Response.Write "<td class='td_detalhe_valor'>"&sAux&"</td>"
					Response.Write "</tr>"&vbcrlf
					conta_informacoes = 1
				end if
				
				'Processa etrada principal
				if xmlArtigo.nodeName  = "ENT_PRINC" then
					sAutor = xmlArtigo.attributes.getNamedItem("Valor").value					
					Response.Write "<tr>"
					Response.Write "<td class='td_detalhe_descricao' style='width: 90px'>&nbsp;"&getTermo(global_idioma, 4109, "Ent. princ.", 0)&"&nbsp;</td>"
					Response.Write "<td class='td_detalhe_valor'>"&sAutor&"</td>"
					Response.Write "</tr>"							
					conta_informacoes = 1					
				end if

				'Processa autores
				if xmlArtigo.nodeName  = "AUTORES" then
					sAux = ""
					For Each xmlAutores In xmlArtigo.childNodes				
						if xmlAutores.nodeName  = "AUTOR" then
							sAutor = xmlAutores.attributes.getNamedItem("Valor").value					
							sSeq   = xmlAutores.attributes.getNamedItem("Seq").value					 
							
							if sAux <> "" then
								sAux = sAux & "<br/>"
							end if
								
							sAux = sAux & sSeq & " " & sAutor
						end if
					Next

					Response.Write "<tr>"
					Response.Write "<td class='td_detalhe_descricao' style='width: 90px'>&nbsp;"&getTermo(global_idioma, 4157, "Ent. sec.", 0)&"&nbsp;</td>"
					Response.Write "<td class='td_detalhe_valor'>"&sAux&"</td>"
					Response.Write "</tr>"							
					conta_informacoes = 1					
				end if

				'Processa resumo
				if xmlArtigo.nodeName  = "RESUMO" then
					sResumo = replace(xmlArtigo.attributes.getNamedItem("Valor").value,Chr(13),"<br/>")
					Response.Write "<tr>"
					Response.Write "<td class='td_detalhe_descricao' style='width: 90px'>&nbsp;"&getTermo(global_idioma, 1214, "Resumo", 0)&"&nbsp;</td>"
					Response.Write "<td class='td_detalhe_valor'>"&sResumo&"</td>"
					Response.Write "</tr>"&vbcrlf
					conta_informacoes = 1										
				end if

				'Verifica se tem mídia
				if xmlArtigo.nodeName  = "MIDIAS" then
					if Veio_Popup = true then
						s_midia = "<a class='link_classic' title='"&getTermo(global_idioma, 1622, "Visualizar mídia", 0)&"' href=""javascript:parent.parent.abrePopup3('asp/midia.asp?tit_artigo="&sTit_Artigo&"&tipo=3&codigo="&iArtigo&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 218, "Mídias", 0)&"',500,490,true,true);"">"&getTermo(global_idioma, 1416, "Clique aqui para ver as mídias do artigo", 0)&"</a>"
					else
						s_midia = "<a class='link_classic' title='"&getTermo(global_idioma, 1622, "Visualizar mídia", 0)&"' href=""javascript:parent.parent.abrePopup2('asp/midia.asp?tit_artigo="&sTit_Artigo&"&tipo=3&codigo="&iArtigo&"&iIndexSrv="&iIndexSrv&"&iBanner="&global_tipo_banner&"&iEscondeMenu="&global_esconde_menu&"&iIdioma="&global_idioma&"','"&getTermo(global_idioma, 218, "Mídias", 0)&"',500,490,true,true);"">"&getTermo(global_idioma, 1416, "Clique aqui para ver as mídias do artigo", 0)&"</a>"
					end if
					Response.Write "<tr>"
					Response.Write "<td class='td_detalhe_descricao' style='width: 90px'>&nbsp;"&getTermo(global_idioma, 218, "Mídias", 0)&"&nbsp;</td>"
					Response.Write "<td class='td_detalhe_valor'>&nbsp;"&s_midia&"&nbsp;</td>"
					Response.Write "</tr>"&vbcrlf
				end if	
			Next										
		end if		
		Response.write "</table>"		
	
		if global_versao = vSOPHIA then
			Response.write "<table class='max_width'>"
			Response.write "<tr><td class='td_detalhe_descricao centro' style='text-align: center; width: auto'>&nbsp;<b>"&getTermo(global_idioma, 792, "Referência bibliográfica", 0)&"</b>&nbsp;</td></tr>"
			Response.write "<tr><td class='centro'><div name='div_refbib' id='div_refbib'></div></td></tr>"
			Response.write "</table>"
			Response.write "<script type='text/javascript'>exibeRefbib_Artigo(3,"&iArtigo&","&iCodObra&","&iCodEx&","&iIndexSrv&")</script>"
			conta_informacoes = 1					
		end if			
		    
        if (iCodObra = 0) then
            descricao_Hint = getTermo(global_idioma, 7676, "Ver detalhes do artigo", 0)
            descricao = getTermo(global_idioma, 7677, "Detalhes do artigo", 0)
        else
            descricao_Hint = getTermo(global_idioma, 7675, "Ver detalhes da analítica", 0)
            descricao = getTermo(global_idioma, 1548, "Detalhes da analítica", 0)

        end if
            		
        class_detalhe = "td_detalhe_valor"
        sTit_fonte = "<a class='link_classic2' title='"&descricao_Hint&"...' href=""javascript:parent.parent.LinkDetalhes(parent.parent.parent.hiddenFrame.modo_busca,1,0,"&iArtigo&",0,'artigo',3);""><span class='transparent-icon span_imagem icon_16 icon-small-search-h' data-icon='search'></span>&nbsp;&nbsp;"&descricao&"</a>"
		Response.Write "<td class='"&class_detalhe&"'>" &sTit_fonte& "&nbsp;</td>"
        
        if conta_informacoes = 0 then
			Response.Write getTermo(global_idioma, 1273, "Nenhuma informação encontrada", 0)
		end if
	end if
%>
</body>
</html>