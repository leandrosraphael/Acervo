<%	 
'--------------------------------------------------------------------------------
' 				MONTA RESULTADO EM FORMATO DE FICHA
'--------------------------------------------------------------------------------
'******************************************************************************************
Set ROService = ROServer.CreateService("Web_Consulta")
sXMLFichas = ROService.MontaFichasSolicitacaoImagem(Ano, Pagina)
Set ROService = nothing
'******************************************************************************************
Response.Write "<table class='max_width' style='border-spacing: 1px; padding: 0'>"

if (left(sXMLFichas,5) = "<?xml") then
	Set xmlDoc = CreateObject("Microsoft.xmldom")
	xmlDoc.async = False
	xmlDoc.loadxml sXMLFichas
	Set xmlRoot = xmlDoc.documentElement
	if xmlRoot.nodeName = "Pagina" then
		For Each xmlFicha In xmlRoot.childNodes
			if xmlFicha.nodeName  = "FICHA_IMAGEM" then
				
				'//-------------------------------------------------------------//
				'---------------------CELULA COM A FICHA------------------------//
				'//-------------------------------------------------------------//
                ficha = ""
				ficha = ficha & "<tr><td class='td_center_top td_grid_ficha_background' style='width: 120px; height: 165px; white-space: nowrap; border-bottom: 1px solid #ccc;'>"
				ficha = ficha & "<table style='display: inline-table'><tr><td>"

                if  (len(trim(xmlFicha.attributes.getNamedItem("IMAGEM_CONCEDIDA").value)) <> 0) then
                    imagem_origem = "<img src=""data:image/"&xmlFicha.attributes.getNamedItem("EXTENSAO").value&";base64,"&xmlFicha.attributes.getNamedItem("IMAGEM_CONCEDIDA").value&""" style='max-width: 100px;'>"
                else
                    imagem_origem =  "<div class='capa_fantasia' style='width: 100px; height: 125px;'></div>"
                end if
    
				ficha = ficha & imagem_origem
				ficha = ficha & "</td></tr></table>"

                ficha = ficha & "<td class='td_center_top td_grid_ficha_background' style='border-bottom: 1px solid #ccc;'>"
				
                ficha = ficha & "<table class='max_width' style='border-spacing: 1px; padding: 0'>"

				ficha = ficha & "<tr>"
				ficha = ficha & "<td class='td_ficha_esq direita'>Solicitação"
				ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&xmlFicha.attributes.getNamedItem("SOLICITACAO").value&"</td></tr>"

				ficha = ficha & "<tr>"
				ficha = ficha & "<td class='td_ficha_esq direita'>Instituição"
				ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&xmlFicha.attributes.getNamedItem("SOLICITANTE_NOME").value&"</td></tr>"

				ficha = ficha & "<tr>"
				ficha = ficha & "<td class='td_ficha_esq direita'>Pessoa Física"
				ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&xmlFicha.attributes.getNamedItem("RESPONSAVEL_NOME").value&"</td></tr>"

				ficha = ficha & "<tr>"
				ficha = ficha & "<td class='td_ficha_esq direita'>Finalidade"
				ficha = ficha & "</td><td class='td_ficha_dir esquerda'>"&xmlFicha.attributes.getNamedItem("DESTINACAO_FINALIDADE").value&"</td></tr>"

				ficha = ficha & "</table></td></tr>"
            
                Response.write ficha
                
            end if
		Next

        Response.write "<script>"
        Response.write " $(""#RegistrosAno"").html(" &xmlRoot.attributes.getNamedItem("TOTALREGISTROS").value&");"
        Response.write " $(""#anoAtual"").html("&Ano&");"
        Response.write "</script>"

	End if
	Set xmlDoc = nothing
	Set xmlRoot = nothing
End if

Response.Write "</table>"
%>