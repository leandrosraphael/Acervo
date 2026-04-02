<%
	sDiretorioArq = "asp"
    nao_imprime_variaveis_globais = "1"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%
    '******************************************************************************************
    AnoAtual = Request.QueryString("AnoAtual")

    Set ROService = ROServer.CreateService("Web_Consulta")
    sXMLFichas = ROService.ObterDadosPaginacaoAno(AnoAtual)
    Set ROService = nothing
    if len(trim(sXMLFichas)) <> 0 then
        if (left(sXMLFichas,5) = "<?xml") then
	
            Set xmlDoc = CreateObject("Microsoft.xmldom")
            xmlDoc.async = False
            xmlDoc.loadxml sXMLFichas
            Set xmlRoot = xmlDoc.documentElement
            if xmlRoot.nodeName = "PAGINACAO" then
                NumPaginas = xmlRoot.attributes.getNamedItem("NumPaginas").value
                TotalRegistros = xmlRoot.attributes.getNamedItem("TotalRegistros").value
	        end if
	        Set xmlDoc = nothing
	        Set xmlRoot = nothing		

	        Response.Write "<script type='text/javascript'>"
	        Response.Write "	parent.hiddenFrame.img_paginas = "&NumPaginas&";"
	        Response.Write "	parent.hiddenFrame.img_total_registros_ano = "&TotalRegistros&";"
	        Response.Write "</script>"
        end if
    end if
    '******************************************************************************************
    
    ListaAnos = cStr(Request.QueryString("listaAnos"))
    IndiceAno = Request.QueryString("indiceAno")
    numAnos = Request.QueryString("numAnos")
    TotalImagens = Request.QueryString("totalImagens")
    PaginaAtual = Request.QueryString("paginaAtual")
    if (NumPaginas = "") then
       NumPaginas = Request.QueryString("numPaginas")
    end if
    if (TotalRegistros = "") then
       TotalRegistros = Request.QueryString("totalRegistros")
    end if
 
    PaginacaoSolicitacaoImagemAno ListaAnos, numAnos, IndiceAno, TotalRegistros, TotalImagens

    PaginacaoSolicitacaoImagemReg NumPaginas, PaginaAtual, TotalRegistros, TotalImagens
%>