<!DOCTYPE HTML>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<% sDiretorioArq="asp" %>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%

%><!-- #include file="../libasp/updChannelProperty.asp" --><%

%>
<title></title>
	<% 
		sDbId = Trim(Request.QueryString("dbid"))
		sAn = Trim(Request.QueryString("an"))
		sTextType = Trim(Request.QueryString("type"))
		if(Session("Logado")= "sim") then
				
			eds_session = Request.Cookies("eds_session")
			Set ROService = ROServer.CreateService("Web_Consulta")
			Set objTextoCompletoEds = ROService.ObterTextoCompletoEds(eds_session, sDbId, sAn, sTextType, global_idioma)
				if objTextoCompletoEds.Result then
					xml = objTextoCompletoEds.sMsg
					texto = ""
					if (left(xml,5) = "<?xml") then 
						Set xmlDoc = CreateObject("Microsoft.xmldom")
						xmlDoc.async = False
						xmlDoc.loadxml xml
						Set xmlRoot = xmlDoc.documentElement
						if xmlRoot.nodeName = "FullText" then
							sSessionToken = ""
							sSessionTimeout = ""
							For Each xmlFullText In xmlRoot.childNodes
								if xmlFullText.nodeName  = "SessionToken" then
									sSessionToken = xmlFicha.text
								end if
								if xmlFullText.nodeName  = "SessionTimeout" then
									sSessionTimeout = xmlFicha.text
								end if
								if xmlFullText.nodeName = "Text" then
									texto = xmlFullText.text
								end if
							Next
						end if
						Set xmlDoc = nothing
						Set xmlRoot = nothing
						if sSessionTimeout <> "" then
							iSessionTimeout = CInt(sSessionTimeout)
							dataExp = DateAdd("n",iSessionTimeout,Now)
							Response.Cookies("eds_session") = sSessionToken 
							Response.Cookies("eds_session").Expires = dataExp
						end if

						Set xmlDoc = nothing
						Set xmlRoot = nothing	
					end if

				end if
		end if
	%>	
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

	<% 
	if sTextType = CStr(EDS_ARQ_PDF) then
	%>
<script type="text/javascript">

	if (parent.parent.hiddenFrame != null) {
		Hiddenfrm = parent.parent.hiddenFrame;
		Hiddenfrm.popup_refresh = true;
		Hiddenfrm.content = 'eds';
	}

</script>
<%
	end if
%>
</head>
<body class="popup" onload="parent.fechaLoadingPopup();">
<table style="width: 100%; text-align: left;">
<tr>
<td style="vertical-align: top">
	
		<% 

			if(Session("Logado")= "sim") then
				
				if texto <> "" then
					if sTextType = CStr(EDS_ARQ_HTML) then
						texto = Replace(texto, "&lt;", "<")
						texto = Replace(texto, "&gt;", ">")
						texto = Replace(texto, "&quot;", """")
						Response.Write "<span style='white-space: pre-line; text-align: left;'>"
						Response.Write texto
						Response.Write "</span>"
					else
						Response.Write "<script type='text/javascript'>parent.fechaLoadingPopup();window.location = '"&texto&"';</script>"
					end if
				end if

			else
				Response.Write "Acesso negado"
			end if

		%>
</td>
</tr>
</table>
</body>
</html>
