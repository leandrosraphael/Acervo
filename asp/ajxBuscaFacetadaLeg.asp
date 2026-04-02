<%
	sDiretorioArq = "asp"
    nao_imprime_variaveis_globais = "1"
%>

<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<%
    iIndexSrv = IntQueryString("iIndexSrv", 1)
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%

    if (Request.Form("alterou") = "1") then
        if (Request.Form("faceta_leg_norma") <> "") then
            Session("Faceta_leg_norma_Srv" & iIndexSrv) = Request.Form("faceta_leg_norma")
        else
            Session("Faceta_leg_norma_Srv" & iIndexSrv) = ""
        end if

        if (Request.Form("faceta_leg_orgao") <> "") then
            Session("Faceta_leg_orgao_Srv" & iIndexSrv) = Request.Form("faceta_leg_orgao")
        else
            Session("Faceta_leg_orgao_Srv" & iIndexSrv) = ""
        end if

        if (Request.Form("faceta_leg_situacao") <> "") then
            Session("Faceta_leg_situacao_Srv" & iIndexSrv) = Request.Form("faceta_leg_situacao")
        else
            Session("Faceta_leg_situacao_Srv" & iIndexSrv) = ""
        end if

        if (Request.Form("faceta_leg_ano") <> "") then
            Session("Faceta_leg_ano_Srv" & iIndexSrv) = Request.Form("faceta_leg_ano")
        else
            Session("Faceta_leg_ano_Srv" & iIndexSrv) = ""
        end if

        if (Request.Form("faceta_leg_assunto") <> "") then
            Session("Faceta_leg_assunto_Srv" & iIndexSrv) = Request.Form("faceta_leg_assunto")
        else
            Session("Faceta_leg_assunto_Srv" & iIndexSrv) = ""
        end if
    end if
	
	On Error Resume Next

	SET ROService = ROServer.CreateService("Web_Consulta")

    XMLFaceta = ROService.MontarFacetaLegislacao(Session("TabTempLeg_Srv" & iIndexSrv), Session("Faceta_leg_norma_Srv" & iIndexSrv), _
        Session("Faceta_leg_orgao_Srv" & iIndexSrv), Session("Faceta_leg_situacao_Srv" & iIndexSrv), _
        Session("Faceta_leg_ano_Srv" & iIndexSrv), Session("Faceta_leg_assunto_Srv" & iIndexSrv), 4, 0, global_idioma)

    if (left(XMLFaceta, 5) = "<?xml") then
	    Set xmlDoc = CreateObject("Microsoft.xmldom")
	    xmlDoc.async = False
	    xmlDoc.loadxml XMLFaceta
   
        Set xmlRoot = xmlDoc.documentElement

        F = 0
        HtmlId = 0
        Facetas = ""
        FacetasSelecionadas = ""

        Response.Write "<form id='faceta-form-srv" & iIndexSrv & "' data-servidor=" & iIndexSrv & " method='POST'>"
	    
        if xmlRoot.nodeName = "facetas" then

            Facetas = "<table class='tab-busca-facetada max_width remover_bordas_padding esquerda'><tr style='height: 26px'><td class='td_faceta_cabecalho'>&nbsp;" & getTermo(global_idioma, 1818, "Filtros", 0) & "</td></tr>"
		    Facetas = Facetas & "<tr><td class='td_facetas'>"

		    For Each xmlFaceta In xmlRoot.childNodes

                idGrupo = xmlFaceta.attributes.getNamedItem("id").value
                nomeFaceta = xmlFaceta.attributes.getNamedItem("nome").value
                quantidadeItens = CLng(xmlFaceta.attributes.getNamedItem("quantidade").value)
                totalItens = CLng(xmlFaceta.attributes.getNamedItem("total").value)
                indiceFaceta = CInt(xmlFaceta.attributes.getNamedItem("indice").value)

                if (quantidadeItens = 1) then
                    style = "display: none;"
                else
                    style = "display: block;"
                    F = F + 1
                End if

                selecaoFaceta = Session("Faceta_leg_" & idGrupo & "_Srv" & iIndexSrv)

                if (F > 1) then 
                    style = style & " margin-top: 20px;"
                end if

                Facetas = Facetas & "<ul class='ul-faceta'" & margin & " style='" & style & "' data-indice='" & indiceFaceta & "'>"
                Facetas = Facetas & "<li class='titulo-faceta expandida'><span>" & nomeFaceta & "</span><span class='transparent-icon span_imagem icon_16 icon-small-down-b'></span></li>"
                For Each xmlItem In xmlFaceta.childNodes
                    idFaceta = xmlItem.attributes.getNamedItem("id").value
                    descFaceta = xmlItem.attributes.getNamedItem("descricao").value
                    qtdeFaceta = xmlItem.attributes.getNamedItem("quantidade").value

                    HtmlId = HtmlId + 1

                    selecionado = ""            
                    If (selecaoFaceta <> "") Then
                        selecaoFacetaArray = Split(selecaoFaceta, ",")
                        For Each valor In selecaoFacetaArray
                            If (ltrim(rtrim(valor)) = idFaceta) Then
                                selecionado = " checked"

                                FacetasSelecionadas = FacetasSelecionadas & "<li class='item-faceta faceta-selecionada'>"
                                FacetasSelecionadas = FacetasSelecionadas & "<a href='#' class='link_custom remover-faceta' data-checkid='faceta-check-" & iIndexSrv & "_" & HtmlId & "'>"
                                FacetasSelecionadas = FacetasSelecionadas & "<span class='transparent-icon span_imagem div_imagem_ajusta div_imagem_right icon_13 icon-small-delete-b'></span>" & descFaceta
                                FacetasSelecionadas = FacetasSelecionadas & "</a></li>"

                                Exit For
                            End If
                        Next
                    End If

                    Facetas = Facetas & "<li class='item-faceta-leg'><input type='checkbox' name='faceta_leg_" & idGrupo & "' id='faceta-check-" & iIndexSrv & "_" & HtmlId & "' value='" & idFaceta & "'" & selecionado & ">" & descFaceta & " (" & qtdeFaceta & ")</li>"
                Next

                if (totalItens > quantidadeItens) then
                    Facetas = Facetas & "<li class='item-faceta-leg ver-mais'><a href='#' class='link_classic2'>" & getTermo(global_idioma, 7046, "Ver mais", 0) & "</a></li>"
                end if

                Facetas = Facetas & "</ul>"
            Next

            if (F = 0) then
                Facetas = Facetas & getTermo(global_idioma, 2024, "Nenhum filtro", 0)
            end if

            Facetas = Facetas & "</td></tr>"
            Facetas = Facetas & "</table>"
        end if

        Set xmlRoot = nothing

        if (FacetasSelecionadas <> "") then
            Response.Write "<table class='tab-busca-facetada max_width remover_bordas_padding esquerda'><tr style='height: 26px'><td class='td_faceta_cabecalho'>&nbsp;" & getTermo(global_idioma, 1716, "Filtros selecionados", 0) & "</td></tr>"
		    Response.Write "<tr><td class='td_facetas'>"
            Response.Write "<ul class='ul-faceta'>"

            Response.Write FacetasSelecionadas

            Response.Write "</ul>"
            Response.Write "</td></tr>"
            Response.Write "</table>"
        end if

        Response.Write Facetas

        Response.Write "</form>"
    end if

    SET ROService = nothing
%>