<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
    <head>
		<title><%=titulo_msg%></title>
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
                
        <link href="../inc/contraste.css" rel="stylesheet" type="text/css" />		
        <link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
		
        <script type="text/javascript">
	    function EmitirCertidao(modo_busca) {
		    mainFrame = parent.parent.mainFrame;
		    parent.fechaPopup();		
		    mainFrame.location =  "../index.asp?modo_busca=" + modo_busca + "&content=certidao_negativa";		
	    }
	    </script>		

	</head>
    <body class="popup" onload="parent.fechaLoadingPopup();">
	    <div class="div_certidao">

		    <table style="height: 84px;" class="remover_bordas_padding max_width">
			    <tr class="centro">
                    <td>&nbsp;&nbsp;<%Response.Write getTermo(global_idioma, 9698, "Confirma a emissão da certidão negativa de débitos?", 0)%></td>
			    </tr>			
			    <tr>
				    <td>
					    <div class="div_certidao_botoes centro">
                            <%                            
                            Response.Write "<input type='button' name='submit' value='"&getTermo(global_idioma, 4, "Confirmar", 0)&"' onclick=EmitirCertidao('"&GetModo_Busca&"');>"
						    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
						    Response.Write "<input type='button' onclick=parent.fechaPopup('asp'); value='"&getTermo(global_idioma, 5, "Cancelar", 0)&"' />"
                            %>
					    </div>
				    </td>
			    </tr>
		    </table>
	    </div>
	</body>
</html>