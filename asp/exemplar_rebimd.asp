<%
    Response.ContentType = "text/html"  
    Response.AddHeader "Content-Type", "text/html;charset=ISO-8859-1"  
    Response.CodePage = 1252  
    Response.CharSet = "ISO-8859-1"

    sDiretorioArq="asp" 
    Response.ContentType = "text/html"
    nao_imprime_variaveis_globais = "1"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%
	iIndexSrv = IntQueryString("servidor", 1)

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	
	codigo_sophia = Request.QueryString("codigo_sophia")

	On Error Resume Next

    Set ROService = ROServer.CreateService("Web_Consulta")
    Set objValida = ROServer.CreateComplexType("TInteiro")
    Set objValida = ROService.ValidaCodigoSophia(codigo_sophia)
    codigo_sophia = objValida.Param0
    tipo_sophia = objValida.Param1
    Set objValida = nothing
    
    TrataErros(1)

    if (cLng(codigo_sophia) <= 0) then  
%>          
        Código sophia inválido.
<%
    else
        xml_exemplar = ROService.MontaExemplar(cLng(codigo_sophia), cInt(tipo_sophia), 0, 0, 0, "#TODOS#", 0, global_idioma, (global_IP_Local = 1))

        TrataErros(1)

        if (trim(xml_exemplar) <> "") then
%>
            <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tamanho_centro">
                <tr>
                    <td align="center" valign="middle">
                        <table width="100" cellpadding="0" border="0" cellspacing="0" class="borda_azul">
                            <tr>
                                <td>
                                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                        <tr>
                                            <td height="10">
                                                <img src="./img/sombra_canto_et.jpg" width="10" height="10">
                                            </td>
                                            <td height="10" background="./img/sombra_top.jpg">
                                                <img src="./img/transp.gif" width="10" height="10">
                                            </td>
                                            <td>
                                                <img src="./img/sombra_canto_dt.jpg" width="10" height="10">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td background="./img/sombra_esq.jpg">
                                                &nbsp;
                                            </td>
                                            <td class="borda_azul">
                                                <table width="100" border="0" cellpadding="0" cellspacing="0" background="./img/fundo.jpg">
                                                    <tr>
                                                        <td height="1" colspan="2" align="center" valign="top" class="det_acesso_usuario">
                                                            <img src="./img/transp.gif" width="10" height="4">
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td width="524" height="30" align="center" valign="middle" class="acesso_usuario">
                                                            <strong>Bibliotecas Integradas do Exército</strong>
                                                        </td>
                                                        <td width="76" align="center" valign="middle" class="acesso_usuario">
                                                            <a href="javascript:MM_showHideLayers('detalhesexemplar','','hide');" class="bt_sair">
                                                                <b>fechar (X)</b></a>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td height="50" colspan="2" align="center" valign="top" bgcolor="#FFFFFF">
                                                            <br />
                                                            <table width="600" height="11%" border="0" align="center" cellpadding="0" cellspacing="0">
                                                                <tr>
                                                                    <td align="center" bgcolor="#FFFFFF">
                                                                        <table width="742" border="0" cellspacing="1" cellpadding="0">
                                                                            <tr nowrap="">
<%
            if left(xml_exemplar,5) = "<?xml" then
	            Set xmlDoc = CreateObject("Microsoft.xmldom")
	            xmlDoc.async = False
	            xmlDoc.loadxml xml_exemplar
	            Set xmlRoot = xmlDoc.documentElement
	            if xmlRoot.nodeName = "FICHA" then
		            For Each xmlPNode In xmlRoot.childNodes
			            '****************************************
			            ' MONTA CABEÇALHO DO GRID DE EXEMPLARES
			            '****************************************
			            if xmlPNode.nodeName = "COLUNAS" then
				            colunas_grid = ""
				            For Each xmlColunas In xmlPNode.childNodes
					            if ((xmlColunas.nodeName = "CODEX") or (xmlColunas.nodeName = "TOMBO") or _
                                    (xmlColunas.nodeName = "EDICAO") or (xmlColunas.nodeName = "ANO") or _
                                    (xmlColunas.nodeName = "VOLUME") or (xmlColunas.nodeName = "NUMERO") or _
                                    (xmlColunas.nodeName = "PARTE") or (xmlColunas.nodeName = "PERC_CIRC") or _
                                    (xmlColunas.nodeName = "SUPORTE") or (xmlColunas.nodeName = "DATA_PUBLICACAO") or _
                                    (xmlColunas.nodeName = "NUM_CHAMADA") or (xmlColunas.nodeName = "CMP_OPCIONAL") or _
                                    (xmlColunas.nodeName = "TAB_OPCIONAL") or (xmlColunas.nodeName = "BIBLIOTECA") or _
                                    (xmlColunas.nodeName = "SITUACAO")) then
						            descricao =  xmlColunas.attributes.getNamedItem("Descricao").value					            
%>
                                                                                <td nowrap="" width="144" align="center" valign="middle" class="tela_exemplar_titulo"
                                                                                    scope="row">
                                                                                    <div class="txt_branco">
                                                                                        <%=descricao%></div>
                                                                                </td>
<%
                                end if

				            Next
			            end if
%>
                                                                            </tr>
<%
			            '****************************************
			            ' GRID DOS EXEMPLARES
			            '****************************************
			            if xmlPNode.nodeName = "EXEMPLARES" then
				            NumExemplares = xmlPNode.attributes.getNamedItem("NUM_EXS").value
				            NumReservas = xmlPNode.attributes.getNamedItem("QTDE_RESERVA").value
								
				            For Each xmlExemplares In xmlPNode.childNodes
					            '****************************************
					            ' CADA EXEMPLAR
					            '****************************************
					            if xmlExemplares.nodeName = "EXEMPLAR" then
%>
                                                                            <tr>
<%
						            For Each xmlCampos In xmlExemplares.childNodes

							            if ((xmlCampos.nodeName = "CODEX") or (xmlCampos.nodeName = "TOMBO") or _
                                            (xmlCampos.nodeName = "EDICAO") or (xmlCampos.nodeName = "ANO") or _ 
                                            (xmlCampos.nodeName = "VOLUME") or (xmlCampos.nodeName = "NUMERO") or _
                                            (xmlCampos.nodeName = "PARTE") or (xmlCampos.nodeName = "PERC_CIRC") or _
                                            (xmlCampos.nodeName = "SUPORTE") or (xmlCampos.nodeName = "DATA_PUBLICACAO") or _
                                            (xmlCampos.nodeName = "NUM_CHAMADA") or (xmlCampos.nodeName = "CMP_OPCIONAL") or _
                                            (xmlCampos.nodeName = "TAB_OPCIONAL") or (xmlCampos.nodeName = "BIBLIOTECA") or _
                                            (xmlCampos.nodeName = "SITUACAO")) then
								            valor = xmlCampos.attributes.getNamedItem("Valor").value
%>
                                                                                <td bgcolor="#FFFFFF" class="tab_2" scope="row">
                                                                                    <div align="center" class="txt">
                                                                                        <%=valor%>
                                                                                        <div></div>
                                                                                    </div>
                                                                                </td>
<%
							            end if
						            Next
					            end if
%>
                                                                            </tr>
<%                                                                                                           				
				            Next
			            end if
		            Next
	            end if
            end if
%>
                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                            <td background="./img/sombra_dir.jpg">
                                                &nbsp;
                                            </td>
                                        </tr>
                                        <tr>
                                            <td width="16">
                                                <img src="./img/sombra_canto_ei.jpg" width="10" height="10">
                                            </td>
                                            <td background="./img/sombra_inf.jpg">
                                                <img src="./img/transp.gif" width="10" height="10">
                                            </td>
                                            <td width="16" height="10">
                                                <img src="./img/sombra_canto_di.jpg" width="10" height="10">
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
<%
        end if    	
    end if
    					   	
	SET ROService = nothing
%>