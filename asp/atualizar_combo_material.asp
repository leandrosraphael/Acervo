<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->
<!-- #include file="../libasp/montar_campos.asp" -->
<%
	idMaterial = Request.QueryString("idMaterial")
	if (Request.QueryString("campo1") <> "") then
		sCampo1 = trim(Request.QueryString("campo1"))
	else
		sCampo1 = ""
	end if
	
	if (Request.QueryString("campo2") <> "") then
		sCampo2 = trim(Request.QueryString("campo2"))
	else
		sCampo2 = ""
	end if
	
	if (Request.QueryString("campo3") <> "") then
		sCampo3 = trim(Request.QueryString("campo3"))
	else
		sCampo3 = ""
	end if
	
	if (Request.QueryString("campo4") <> "") then
		sCampo4 = trim(Request.QueryString("campo4"))
	else
		sCampo4 = ""
	end if

	if (Request.QueryString("campo5") <> "") then
		sCampo5 = trim(Request.QueryString("campo5"))
	else
		sCampo5 = ""
	end if
	
	if (Request.QueryString("campo6") <> "") then
		sCampo6 = trim(Request.QueryString("campo6"))
	else
		sCampo6 = ""
	end if
	
	if (Request.QueryString("campo7") <> "") then
		sCampo7 = trim(Request.QueryString("campo7"))
	else
		sCampo7 = ""
	end if
	
	if (Request.QueryString("campo8") <> "") then
		sCampo8 = trim(Request.QueryString("campo8"))
	else
		sCampo8 = ""
	end if

	On Error Resume Next
	xmlCamposMaterial = ROService.GetCamposMaterial(idMaterial)
    sMsgErro = TrataErros
    
    if (sMsgErro <> "") then
        Response.Write sMsgErro
    else
        xmlCamposMaterial = replace(xmlCamposMaterial, "windows-1252", "utf-8")
        response.Write FormataCamposBusca(xmlCamposMaterial,sCampo1,sCampo2,sCampo3,sCampo4,sCampo5,sCampo6,sCampo7,sCampo8, "5")
    end if
%>