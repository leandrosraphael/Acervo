<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
	<title></title>
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
<% end if %>
<link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
</head>
<body class="bibcomp_body">
<%
iIndexSrv = IntQueryString("servidor", 1)
'O índice iIndexSrv que define em qual servidor será acessado
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

tipo = trim(Request.querystring("tipo"))
cod = trim(request.QueryString("cod"))

On Error Resume Next
SET ROService = ROServer.CreateService("Web_Consulta")
xml_bibcomp = ROService.BibComplementar(cod,global_idioma)
SET ROService = nothing
TrataErros(1)

Response.Write Formata_BibComp(tipo,cod,xml_bibcomp)

Function Formata_BibComp(tipo,cod,stringXML)	
	if left(stringXML,5) <> "<?xml" then
		Formata_BibComp = ""
	else
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml stringXML
		Set xmlRoot = xmlDoc.documentElement
		strDoc = "<table style='padding: 0; border-spacing: 0; width: 100%'>"
		strDoc = strDoc & "<tr><td class='descritor'>"&getTermo(global_idioma, 25, "Descrição", 0)&"</td></tr>"
		For Each xmlPNode In xmlRoot.childNodes
			If xmlRoot.childNodes.length = 0 Then
				strDoc = ""				
			Else
				if xmlPNode.nodeName = tipo then
					For Each xmlNode In xmlPNode.childNodes
						if tipo = "Autores" then
							nome_funcao = "LinkBuscaAutor"
						else
							nome_funcao = "LinkBuscaAssunto"
						end if
						if xmlNode.nodeName  = "REG" then
							Descricao = trim(xmlNode.attributes.getNamedItem("Descricao").value)
							Codigo = trim(xmlNode.attributes.getNamedItem("Codigo").value)
							desc_formatada = replace(Descricao," ","_")
							strDoc = strDoc & "<tr><td class='bibcomp'><a class='link_bibcomp' href='#bibComp' onClick="&nome_funcao&"('iframe',"&cod&",'"&desc_formatada&"',"&iIndexSrv&")>" &Descricao&"</a></td></tr>"
						end if									
					next
				end if				
			End If
		Next
		strDoc = strDoc & "</table>"		
		Set xmlDoc = nothing
		Set xmlRoot = nothing
		Formata_BibComp = strDoc
	End If
End Function

%>
</body>
</html>