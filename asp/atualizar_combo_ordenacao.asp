<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->
<%
	codigoObjeto = Request.QueryString("idMaterial")
	codigoCampoOrdenacao = Request.QueryString("campo_ordenacao")
	codigoUsuario = ClNg(Session("codigo"))

	On Error Resume Next
	xmlCamposOrdenacaoObjeto = ROService.ObterListaCamposOrdenacao(codigoObjeto, codigoUsuario)
    sMsgErro = TrataErros
    
    if (sMsgErro <> "") then
        select_html = sMsgErro
    else
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml xmlCamposOrdenacaoObjeto
		Set xmlRoot = xmlDoc.documentElement

		if xmlRoot.nodeName  = "CAMPOS" then
			select_html =	"<td class='direita'>Ordenação</td>" & _
							"<td class='esquerda'>"

			select_html =  select_html & _
				"<select name='campo_ordenacao'>"

			For Each xmlCampo In xmlRoot.childNodes
				if xmlCampo.nodeName = "CAMPO" then
					select_html = select_html & "<option value='"&xmlCampo.attributes.getNamedItem("CODIGO").value & "'"

					if (CInt(xmlCampo.attributes.getNamedItem("CODIGO").value) = CInt(codigoCampoOrdenacao)) then
						select_html = select_html & " selected"
					end if

					select_html = select_html & ">" & xmlCampo.attributes.getNamedItem("NOME").value&"</option>"
				end if
			next

			select_html = select_html & "</select></td>"
		end if
		
		Set xmlRoot = nothing
		Set xmlDoc = nothing
	end if

	response.write select_html
%>