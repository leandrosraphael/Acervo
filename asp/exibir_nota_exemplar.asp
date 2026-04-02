<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html style="height: 100%;" <%=htmlClass%>>
<head>
<%
sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />

<% if config_css_estatico = 1 then %>
	<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
<% else %>
	<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
<% end if 
    
    iCodEx = Request.QueryString("exemplar")
	iIndexSrv = IntQueryString("Servidor", 1)
%>
	<title><%=getTermo(global_idioma, 183, "Notas", 0)%></title>
<% 

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 	
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
%>
<script src="../scripts/jquery/jquery-3.2.1.min.js"></script>

	<script type="text/javascript" src="../scripts/jquery/jquery-ui.min.js"></script>
	<script type="text/javascript" src="../scripts/jquery.multiselect.js"></script>

	<link href="../scripts/css/jquery-ui-1.8.18.datepicker.css" rel="stylesheet" />
	<link href="../scripts/css/jquery.multiselect.css" rel="stylesheet" />
	<link href="../scripts/css/multiselect-custom.css" rel="stylesheet" />
    <link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
    <link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />

</head>
<body class="popup" style="height: 80%;" onload="parent.fechaLoadingPopup();">
	<div class="centro" style="width: 100%; height: 100%; margin-top: 20px;">
		<div style="font-weight: bold; margin-bottom: 30px">
			<%=getTermo(global_idioma, 9237, "Notas do exemplar" , 0)%>
		</div>

		<% 
		
		Set ROService = ROServer.CreateService("Web_Consulta")
		notasExemplar = ROService.BuscarNotasExemplar(iCodEx, global_idioma)

		%>

		<div>
			<%
			if (trim(notasExemplar) <> "") then 
				Set xmlNotas = CreateObject("Microsoft.xmldom")
				
				xmlNotas.async = False
				xmlNotas.loadxml notasExemplar
			
				Set notasExemplar = nothing
			
				Set xmlRoot = xmlNotas.documentElement
	
				if (xmlRoot.nodeName = "NOTAS_EXEMPLAR") then
					%>

					<table class="table-nota-exemplar td_grid_ficha_background">
						<tbody>

					<%
					For Each xmlNota In xmlRoot.childNodes
		
						if (xmlNota.nodeName = "NOTA_EXEMPLAR") then
					
							sTituloNota	= xmlNota.attributes.getNamedItem("TITULO_NOTA").Value
							sNota = xmlNota.attributes.getNamedItem("NOTA").value
                            

							response.write "<tr>"
							response.write "	<td class='td_ficha_esq direita'>" & sTituloNota & "</td>"
							response.write "	<td class='td_ficha_dir esquerda'>" & TrocaTagEspaco(sNota) & "</td>"
							response.write "</tr>"
						end if

					Next

					%>

						</tbody>
					</table>

					<%

				end if	
	
				Set xmlNotas = nothing
				Set ROService = nothing
			end if

			%>
		</div>
	</div>
</body>
</html>