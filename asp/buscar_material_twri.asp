<%
On Error Resume Next
Set ROService = ROServer.CreateService("Web_Consulta")
sXML = ROService.BuscarListaMaterialTWRI

Set xmlListaMaterial = CreateObject("Microsoft.xmldom")
xmlListaMaterial.async = False
xmlListaMaterial.loadxml sXML
		
Set xmlRoot = xmlListaMaterial.documentElement
	
if (xmlRoot.nodeName = "Materiais") then

	iQtdMaterial = xmlRoot.attributes.getNamedItem("QTDE").Value

	sListaMaterial = ""

	For Each xmlMaterial In xmlRoot.childNodes

		if (xmlMaterial.nodeName = "Material") then
			sCodigo	=	xmlMaterial.attributes.getNamedItem("CODIGO").value
			sDescricao	=	xmlMaterial.attributes.getNamedItem("DESCRICAO").Value
		end if

		if sCodigo = null or sCodigo = "" or sCodigo = "UNDEFINED" then
			sCodigo = "0"
		end if

		if sNome = "UNDEFINED" then
			sDescricao = ""
		end if

		sListaMaterial = sListaMaterial & _
			"<option value='"&sCodigo&"'>"&sDescricao&"</option>"
	Next
end if

Set xmlListaMaterial = nothing
Set ROService = nothing
%>