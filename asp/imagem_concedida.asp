<%

Ano = Request.QueryString("ano")
'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

if (Ano = "") then
    '******************************************************************************************
    Set ROService = ROServer.CreateService("Web_Consulta")
    sXMLFichas = ROService.ObterDadosPaginacaoSolicitacaoImagem
    Set ROService = nothing
    '******************************************************************************************

    if len(trim(sXMLFichas)) <> 0 then
       if (left(sXMLFichas,5) = "<?xml") then
	
            Set xmlDoc = CreateObject("Microsoft.xmldom")
            xmlDoc.async = False
            xmlDoc.loadxml sXMLFichas
            Set xmlRoot = xmlDoc.documentElement
            if xmlRoot.nodeName = "PAGINACAO" then
	            Anos = xmlRoot.attributes.getNamedItem("Anos").value
                NumAnos = xmlRoot.attributes.getNamedItem("NumAnos").value
                NumPaginas = xmlRoot.attributes.getNamedItem("NumPaginas").value
                TotalRegistros = xmlRoot.attributes.getNamedItem("TotalRegistros").value
                TotalImagens = xmlRoot.attributes.getNamedItem("TotalImagens").value
	        end if
	        Set xmlDoc = nothing
	        Set xmlRoot = nothing		

	        Response.Write "<script type='text/javascript'>"
	        Response.Write "	parent.hiddenFrame.img_vetor_ano = ["&Anos&"];"
	        Response.Write "	parent.hiddenFrame.idx_vetor_ano = "&NumAnos-1&";"
	        Response.Write "	parent.hiddenFrame.img_paginas = "&NumPaginas&";"
	        Response.Write "	parent.hiddenFrame.img_pagina_atual = 1;"
	        Response.Write "	parent.hiddenFrame.img_total_registros_ano = "&TotalRegistros&";"
	        Response.Write "	parent.hiddenFrame.img_total_imagens = "&TotalImagens&";"
	        Response.Write "</script>"
 
            Response.Write "<table id='paginacao-imagem-topo' class='tab_paginacao max_width remover_bordas_padding'>"
            Response.Write "<tr style='height: 26px'><td>"

            '-----------------------------------------------------------------------------------
            PaginacaoSolicitacaoImagemAno Anos, numAnos, numAnos-1, TotalRegistros, TotalImagens
            '-----------------------------------------------------------------------------------

            '------------------------------------------------------------------------
            PaginacaoSolicitacaoImagemReg NumPaginas, 1, TotalRegistros, TotalImagens
            '------------------------------------------------------------------------
            Response.Write "</td>"
            Response.Write "</tr></table>"
            Response.write "<div id='div-ficha-imagem'></div>"

            Response.Write "<table id='paginacao-imagem-rodape' class='tab_paginacao max_width remover_bordas_padding'>"
            Response.Write "<tr style='height: 26px'><td>"
            '-----------------------------------------------------------------------------------
            PaginacaoSolicitacaoImagemAno Anos, numAnos, numAnos-1, TotalRegistros, TotalImagens
            '-----------------------------------------------------------------------------------

            '------------------------------------------------------------------------
            PaginacaoSolicitacaoImagemReg NumPaginas, 1, TotalRegistros, TotalImagens
            '------------------------------------------------------------------------
            Response.Write "</td>"
            Response.Write "</tr></table>"

            Pagina = 1
            arrTemp = Split(Anos, ",")
            Ano = arrTemp(UBound(arrTemp))
            set arrTemp = nothing

            Response.write "<script>"
            Response.write "    ListarFichaImagens("&Ano&","&Pagina&");"
            Response.write "</script>"

	    else
		    Response.Write "	<table class='tab_paginacao max_width remover_bordas_padding'>"
		    Response.Write "	<tr style='height: 26px'>"
		    Response.Write "		<td>"		
		    Response.Write "			<p class='centro'><b>" & getTermo(global_idioma, 1341, "Nenhum registro encontrado", 0) & "</b></p>"
		    Response.Write "		</td>"
		    Response.Write "	</tr>"
		    Response.Write "	</table>"
            Response.write "<div id='div-ficha-imagem'></div>"
	    end if
    else
	    Response.Write "<table class='tab_paginacao max_width remover_bordas_padding'>"
	    Response.Write "<tr style='height: 26px'>"
	    Response.Write "	<td>"
	    Response.Write "		<p class='centro'><b>" & getTermo(global_idioma, 1341, "Nenhum registro encontrado", 0) & "</b></p>"
	    Response.Write "	</td>"
	    Response.Write "</tr>"
	    Response.Write "</table>"
        Response.write "<div id='div-ficha-imagem'></div>"		
    end if
else
    Pagina = 1
    Response.write "<script>"
    Response.write "    ListarFichaImagens("&Ano&","&Pagina&");"
    Response.write "</script>"
end if
%>