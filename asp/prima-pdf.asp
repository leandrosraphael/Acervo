<!DOCTYPE html>

<html>
<head>

<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
    
<%
    codigoMidia = Request.QueryString("codigoMidia")
    nomeArquivo = Request.QueryString("nomeArquivo")
    usuario = Session("codigo")

    On Error Resume Next

	'******************************************
	'Contagem de acesso
	'******************************************
	Set ROService = ROServer.CreateService("Web_Consulta")
	ROService.ContaAcessosMidias CLng(codigoMidia), CLng(usuario)
	Set ROService = nothing
%>

		<title>Prima PDF</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		
        <link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
		<link rel="stylesheet" href="../scripts/css/number-polyfill.css" />
		<link rel="stylesheet" href="../inc/prima-pdf.css" />

        <% if (config_customizacao_layout = 0) then %>
		    <link href="../inc/estilo.css" rel="stylesheet" type="text/css">
	    <% else %>
		    <link href="../libasp/estilo.asp" rel="stylesheet" type="text/css">
	    <% end if %>

	</head>
	<body>
       
		<div class="pr-pdf">
			<div class="pr-pdf-header">
				<div class="pr-pdf-header-pagination">
                    <div>
                        <label>Página</label>
					    <input type="number" id="txtPage" value="1" min="1" />
                        <label>de <span id='spanTotalPages'></span></label>
                    </div>
					<div>
						<div title="Mostrar página anterior" class="transparent-icon pdf-control-prev"></div>
						<div title="Mostrar próxima página" class="transparent-icon pdf-control-next"></div>
					</div>
				</div>
				<div class="pr-pdf pr-pdf-header pr-pdf-header-name"></div>
				<div class="pr-pdf-header-controls">
					<label>Zoom</label>
					<div title="Aumentar zoom" class="transparent-icon pdf-control-zoomIn"></div>
                    <div title="Diminuir zoom" class="transparent-icon pdf-control-zoomOut"></div>
					<div title="Ajustar página à janela" class="transparent-icon pdf-control-expand"></div>
				</div>
			</div>
			<div class="pr-pdf-body">
				<div class="pdf-wrapper">
					<div class="container">
                        <img alt="" src="" />
					</div>					
					<div title="Mostrar página anterior" class="transparent-icon pdf-control-prev side"></div>
					<div title="Mostrar próxima página" class="transparent-icon pdf-control-next side"></div>
				</div>
			</div>
		</div>
        <script src="../scripts/jquery-1.7.1.min.js"></script>
        <script src="../scripts/number-polyfill.min.js"></script>
        <script src="../scripts/prima-pdf.js"></script>
        <script src="../scripts/e-smart-zoom-jquery.js"></script>
		<!--[if lt IE 9]>
			<script src="//html5shiv.googlecode.com/svn/trunk/html5.js"></script>
		<![endif]-->
		<script type="text/javascript">
		    var codigoMidia = "<%=CodigoMidia%>";
		    var nomeArquivo = "<%=NomeArquivo%>";
		    $(function () {
		        var primaPDF = new PrimaPDF();
		        primaPDF.init(nomeArquivo, codigoMidia);
		    });
		</script>
	</body>
</html>