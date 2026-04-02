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
        
        <link href="../inc/avaliacao_votar.css" rel="stylesheet" type="text/css" />
        <link href="../inc/contraste.css" rel="stylesheet" type="text/css" />		
        <link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
		
        <%
		
		iCodigo  = request.QueryString("Codigo")
		iCodigoUsuario = request.QueryString("CodigoUsuario")

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

		<title></title>

	</head>

     <body class="popup" onload="parent.fechaLoadingPopup();">
	    
		<div class="divPrincipal">
		
			<b><%=titulo%></b>

			<% Response.Write "<form name='frm_avaliacao_votar' method='POST' action='avaliacao_processar_voto.asp?iIdioma=" & global_idioma & "&iBanner=" & global_tipo_banner & "&CodigoObra=" & iCodigo &  "&CodigoUsuario=" & iCodigoUsuario & "'>" %>

				<div class="divPrincipal">
			
				<%
					Dim i
					For i=1 to 4
					Response.Write "<p><span class='opcaoVoto' style='margin-right: 10px'><input type='radio' name='nota_avaliacao' value='" & i &  "' /></span><span class='opcaoVoto'>" & ImprimeEstrelas(i & ",0",true) & "</span></p>"
					Next 			

					Response.Write "<p><span class='opcaoVoto' style='margin-right: 10px'><input type='radio' name='nota_avaliacao' value='5' CHECKED/></span><span class='opcaoVoto'>" & ImprimeEstrelas("5,0",true) & "</span></p>"
				 %> 

				 <div id="botoes">
				 
					 <input type="submit" name='confirmar' value="<% Response.Write getTermo(global_idioma, 4, "Confirmar", 0) %>"/>
					 <input type="button" name='cancelar' value="<% Response.Write getTermo(global_idioma, 5, "Cancelar", 0) %>" onclick="javascript:parent.fechaPopup('asp');"/>

				 </div>

				</div>

			</form>

		</div>

    </body>

</html>