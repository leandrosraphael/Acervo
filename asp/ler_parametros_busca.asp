<%
	sDados = trim(Request.QueryString("dados"))
	
	codItem = Request.QueryString("codigo")
	
	if (trim(Request.QueryString("objeto")) <> "") then
		iObjeto = CInt(Request.QueryString("objeto"))
	else
		iObjeto = 0
	end if
	
	if (trim(Request.QueryString("contexto")) <> "") then
		iContexto = CInt(Request.QueryString("contexto"))
	else
		iContexto = 0
	end if

	if (trim(Request.QueryString("pagina")) <> "") then
		iPagina = CInt(Request.QueryString("pagina"))
	else
		iPagina = 1
	end if
	
	if (trim(Request.QueryString("imagem")) <> "") then
		iImagem = CInt(Request.QueryString("imagem"))
	else
		iImagem = 0
	end if
	
	if (trim(Request.QueryString("campo1")) <> "") then
		sCampo1 = trim(Request.QueryString("campo1"))
	else
		sCampo1 = ""
	end if

	if (trim(Request.QueryString("campo2")) <> "") then
		sCampo2 = trim(Request.QueryString("campo2"))
	else
		sCampo2 = ""
	end if
	
	if (trim(Request.QueryString("campo3")) <> "") then
		sCampo3 = trim(Request.QueryString("campo3"))
	else
		sCampo3 = ""
	end if
	
	if (trim(Request.QueryString("campo4")) <> "") then
		sCampo4 = trim(Request.QueryString("campo4"))
	else
		sCampo4 = ""
	end if

	if (trim(Request.QueryString("campo5")) <> "") then
		sCampo5 = trim(Request.QueryString("campo5"))
	else
		sCampo5 = ""
	end if

	if (trim(Request.QueryString("campo6")) <> "") then
		sCampo6 = trim(Request.QueryString("campo6"))
	else
		sCampo6 = ""
	end if
	
	if (trim(Request.QueryString("campo7")) <> "") then
		sCampo7 = trim(Request.QueryString("campo7"))
	else
		sCampo7 = ""
	end if
	
	if (trim(Request.QueryString("campo8")) <> "") then
		sCampo8 = trim(Request.QueryString("campo8"))
	else
		sCampo8 = ""
	end if
	
	sTabTmp = trim(Request.QueryString("tmp"))
	
	if (trim(Request.QueryString("tmp_objeto")) <> "") then
		iTmpObjeto = CInt(Request.QueryString("tmp_objeto"))
	else
		iTmpObjeto = 0
	end if
	
	if (trim(Request.QueryString("tmp_contexto")) <> "") then
		iTmpContexto = CInt(Request.QueryString("tmp_contexto"))
	else
		iTmpContexto = 0
	end if

	if (trim(Request.QueryString("campo_ordenacao")) <> "") then
		CodigoCampoOrdenacao = CLng(Request.QueryString("campo_ordenacao"))
	else
		CodigoCampoOrdenacao = 0
	end if

	if (trim(Request.QueryString("campo_ordem")) <> "") then
		OrdemPesquisa = CInt(Request.QueryString("campo_ordem"))
	else
		OrdemPesquisa = 0
	end if
%>