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
iIdioma = IntQueryString("iIdioma", 0)
veio_de = "servicos"

'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%>
<!-- #include file="../libasp/updChannelProperty.asp" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
</head>
<body><%
	Set ROService = ROServer.CreateService("Web_Consulta")
	sXMLFichas = ROService.MontarListaLevantamentoBibliografico(True, global_idioma)
	Set ROService = nothing
	'******************************************************************************************
		
	if len(trim(sXMLFichas)) <> 0 then
			
        if (left(sXMLFichas,5) = "<?xml") then
	        Set xmlDoc = CreateObject("Microsoft.xmldom")
	        xmlDoc.async = False
	        xmlDoc.loadxml sXMLFichas
	        Set xmlRoot = xmlDoc.documentElement
            nQuantidade = CLng(xmlRoot.attributes.getNamedItem("QUANTIDADE").value)
            if (nQuantidade > 0) then
	            if xmlRoot.nodeName = "LEVANTAMENTO" then
	                Response.Write "<td id='lb-td' class='tr-lev-bib'>"
	                Response.Write "    <div class='lb-conteudo'>"
	                Response.Write "        <div class='lb-cabecalho'>"
	                Response.Write "            <span class='mensagens_titulo'>"&getTermo(iIdioma, 8649, "Levantamentos bibliográficos", 0)&"</span>"
                    Response.Write "        </div>"
	                Response.Write "        <div class='lb-conteudo-resumo'>"
                
                    nReg = 0
                    For Each xmlFicha In xmlRoot.childNodes
                        if xmlFicha.nodeName  = "LEVANTAMENTO_BIB" then
                            nReg = nReg + 1
                            sCodigo = xmlFicha.attributes.getNamedItem("CODIGO").value
                            sAssunto = xmlFicha.attributes.getNamedItem("ASSUNTO").value
                            sDescricaoOriginal  = replace(replace(xmlFicha.attributes.getNamedItem("DESCRICAO").value,Chr(10),"<br/>"),chr(13),"")
						    if (Len(sAssunto) > 47) then
							    sAssunto = left(sAssunto,47) & "..."
						    end if
						    if (Len(sDescricaoOriginal) >= 87) then
                                sLink = "<div>"
							    sDescricao = left(sDescricaoOriginal,87) & "...&nbsp;&nbsp;<span class='span_imagem div_imagem_right_3 icon_9 mais-png '></span><a class='link_classic2' title='"&getTermo(global_idioma, 1605, "Exibir texto completo", 0)&"..."&"' onClick=""ExibirTextoCompleto("&sCodigo&")"" href=""javascript:ExibirTextoCompleto("&sCodigo&");"">"&getTermo(global_idioma, 1621, "Ler mais", 0)&"</a>"
                                if (nReg = 4) and (nQuantidade > 4) then
                                    sDescricao = sDescricao & "<div style='float: right;'><span class='link-ver-mais transparent-icon span_imagem icon_16 icon-small-baloonadd-w'></span><a class='link_serv' href=javascript:LinkLevantamentoBibliografico();>&nbsp;"& getTermo(global_idioma, 7046, "Ver mais", 0) &"</a></div>"
                                end if
                                sLink = sLink & "</div>"
                            else
                                sDescricao = sDescricaoOriginal
                                if (nReg = 4) and (nQuantidade > 4) then
                                    sLink = "<div style='float: right;'><span class='link-ver-mais transparent-icon span_imagem icon_16 icon-small-baloonadd-w'></span><a class='link_serv' href=javascript:LinkLevantamentoBibliografico();>&nbsp;"& getTermo(global_idioma, 7046, "Ver mais", 0) &"</a></div>"
                                else
                                    sLink = ""
                                end if
						    end if
                        
                            Response.Write "<div>"
                            Response.Write "   <a href=javascript:buscarTituloLevantamentoBib("&iIndexSrv&","&sCodigo&");>"
                            Response.Write "      <span style='text-decoration: none;'><p style='text-align: center;'>"&sAssunto&"</p></span></a>"
                            Response.Write "   <br />"
                            Response.Write "  <div>"
                            Response.Write "     <p class='justificado' style='height: 81px; color: #555'>"&sDescricao&"</p>"
                            response.write "     <div>"&sLink&"<div class='clear:both;'></div></div>"
                            Response.Write "  </div>"
                            Response.Write "</div>"
                        end if
                    next
	                Response.Write "        </div>"
	                Response.Write "    </div>"
	                Response.Write "</td>"
                end if   
	            Set xmlDoc = nothing
	            Set xmlRoot = nothing
            End if
            Response.Write "</table>"
			
		    Response.Write "</tr>"
		    Response.Write "</table>"
        end if
	else
	    Response.Write ""
	end if
 %>
</body>
</html>
