<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
    <%
    sDiretorioArq = "asp"
    %>
    <!-- #include file="../config.asp" -->
    <!-- #include file="../idiomas/idiomas.asp" -->
    <!-- #include file="../libasp/header.asp" -->
    <!-- #include file="../libasp/funcoes.asp" -->
    <title></title>
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
	<script src="../scripts/jquery/jquery-3.2.1.min.js"></script>

    <script type="text/javascript">
	    function aplicarSelecao() {
	        var form = $('#faceta-form-ver-mais');
	        var servidor = $(form).attr('data-servidor');
	        var idFaceta = $(form).attr('data-faceta-id');
	        var faceta = $(form).serialize();
	        
            
	        parent.AplicarFacetaVerMais(servidor, idFaceta, faceta);
	        parent.fechaPopup();

		    return true;
	    }
    </script>
</head>
<%
	iIndexSrv = IntQueryString("iIndexSrv", 1)

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%

    iFaceta = IntQueryString("iFaceta", 1)

    On Error Resume Next

	SET ROService = ROServer.CreateService("Web_Consulta")

    XMLFaceta = ROService.MontarFaceta(Session("TabTemp_Srv" & iIndexSrv), Session("Faceta_material_Srv" & iIndexSrv), _
        Session("Faceta_idioma_Srv" & iIndexSrv), Session("Faceta_edicao_Srv" & iIndexSrv), Session("Faceta_ano_Srv" & iIndexSrv), _
        Session("Faceta_autor_Srv" & iIndexSrv), Session("Faceta_assunto_Srv" & iIndexSrv), 0, iFaceta, global_idioma)

    SET ROService = nothing

    sSelecaoFaceta = ""

    Select case iFaceta
        Case 1
            sSelecaoFaceta = Session("Faceta_material_Srv" & iIndexSrv)
        Case 2
            sSelecaoFaceta = Session("Faceta_idioma_Srv" & iIndexSrv)
        Case 3
            sSelecaoFaceta = Session("Faceta_edicao_Srv" & iIndexSrv)
        Case 4
            sSelecaoFaceta = Session("Faceta_ano_Srv" & iIndexSrv)
        Case 5
            sSelecaoFaceta = Session("Faceta_autor_Srv" & iIndexSrv)
        Case 6
            sSelecaoFaceta = Session("Faceta_assunto_Srv" & iIndexSrv)
    End Select
%>
<body class="popup">
    <script type="text/javascript">parent.fechaLoadingPopup();</script>
    <table class="max_width centro">
        <tr>
            <td class="td_center_top">
                <div style="height: 420px; max-height: 420px; overflow: auto;">
                    <br />
                    <%
					expirouSessao = false

                    if (left(XMLFaceta, 5) = "<?xml") then
	                    Set xmlDoc = CreateObject("Microsoft.xmldom")
	                    xmlDoc.async = False
	                    xmlDoc.loadxml XMLFaceta
   
                        Set xmlRoot = xmlDoc.documentElement
    
                        if xmlRoot.nodeName = "facetas" then

		                    For Each xmlFaceta In xmlRoot.childNodes

                            idGrupo = xmlFaceta.attributes.getNamedItem("id").value
                            nomeFaceta = xmlFaceta.attributes.getNamedItem("nome").value
                            quantidadeItens = CLng(xmlFaceta.attributes.getNamedItem("quantidade").value)
                            totalItens = CLng(xmlFaceta.attributes.getNamedItem("total").value)

                            Response.Write "<form id='faceta-form-ver-mais' data-servidor='" & iIndexSrv & "' data-faceta-id='faceta_" & idGrupo & "' method='POST'>"

                            Response.Write "<table class='tab_borda_selecao' style='display: inline-table'><tr><td>"
                            Response.Write "<table class='max_width' style='border-spacing: 1px; padding: 0'>"
                            Response.Write "<tr style='height: 20px'>"

                            Response.Write "<td class='td_tabelas_titulo centro' colspan='2'>"
                            Response.Write nomeFaceta
                            Response.Write "</td>"
                            Response.Write "</tr>"

                            iSeq = 1

                            For Each xmlItem In xmlFaceta.childNodes
                                idFaceta = xmlItem.attributes.getNamedItem("id").value
                                descFaceta = xmlItem.attributes.getNamedItem("descricao").value
                                qtdeFaceta = xmlItem.attributes.getNamedItem("quantidade").value

                                selecionado = ""            
                                If (sSelecaoFaceta <> "") Then
                                    selecaoFacetaArray = Split(sSelecaoFaceta, ",")
                                    For Each valor In selecaoFacetaArray
                                        If (ltrim(rtrim(valor)) = idFaceta) Then
                                            selecionado = " checked"
                                            Exit For
                                        End If
                                    Next
                                End If

                                if (iSeq mod 2) > 0 then '### IMPAR
                                    td_class = "td_tabelas_valor2"
                                else '### PAR
                                    td_class = "td_tabelas_valor1"
                                end if

                                Response.Write "<tr style='height: 20px'>"

                                Response.Write "<td class='"&td_class&" centro' style='width: 35px'>"
                                Response.Write "<input type='checkbox' name='faceta_" & idGrupo & "' value='" & idFaceta & "' " & selecionado & " />"
                                Response.Write "</td>"

                                Response.Write "<td class='" & td_class & " esquerda' style='width: 270px'>&nbsp;"
                                Response.Write descFaceta & " (" & qtdeFaceta & ")"
                                Response.Write "</td>"

                                Response.Write "</tr>"

                                iSeq = iSeq + 1
                            Next

                            Response.Write "</table>"
                            Response.Write "</td></tr></table>"

                            Response.Write "</form>"
                        Next
                        end if

                        Set xmlRoot = nothing
                        Set xmlDoc = nothing
					else
						expirouSessao = true
						Response.Write "<span class='span_imagem div_imagem_right icon_13 alerta0'></span>&nbsp;<span style='color: #006699'><br />"&getTermo(global_idioma, 1384, "Esta sessão expirou.", 0)&" "&getTermo(global_idioma, 1385, "Por favor, refaça sua busca.", 0)&"</span>"
                    end if %>
                </div>
            </td>
        </tr>
        <tr style="height:30px">
            <td>
                <div class="centro">
					<% if (expirouSessao = false) then %>
						<input type='button' value='<%=getTermo(global_idioma, 128, "Selecionar", 0)%>' onclick='return aplicarSelecao();' />
						&nbsp;&nbsp;
					<% end if %>
                    <input type='button' value='<%=getTermo(global_idioma, 5, "Cancelar", 0)%>' onclick='parent.fechaPopup()' />
                </div>
            </td>
        </tr>
    </table>
</body>
</html>