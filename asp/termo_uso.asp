<!-- #include file="../libasp/funcoes.asp" -->
<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->
<% 
if (sMsgErro <> "") then
	Response.Write sMsgErro
else	
	xmlTermoUso = ROService.ObterTermoUso
	sMsgErro = TrataErros
	if (sMsgErro <> "") then
		Response.Write sMsgErro
	else
		if xmlTermoUso <> "" then
			Set xmlDoc = CreateObject("Microsoft.xmldom")
			xmlDoc.async = False
			xmlDoc.loadxml xmlTermoUso
			Set xmlRoot = xmlDoc.documentElement
	
			if xmlRoot.nodeName  = "TERMO_USO" then
	
				For Each xmlPNode In xmlRoot.childNodes
					if xmlPNode.nodeName = "TITULO" then
						titulo  = xmlPNode.attributes.getNamedItem("VALOR").value
					elseif xmlPNode.nodeName = "DESCRICAO" then
						conteudo = xmlPNode.attributes.getNamedItem("VALOR").value			
					end if
				Next	

				if titulo <> "" then
					Response.Write "<div id='divTituloTermoUso'>"
					Response.Write "<span>"&titulo&"</span>"
					Response.Write "</div>"
				end if
		
				Response.Write "<table id='termouso' class='grid'>"
				Response.Write "<tr>"		
				Response.Write "<td class='td_valor'>"&conteudo&"</td>"
				Response.Write "</tr>"		
				Response.Write "</table>"		
			end if
	
			Set xmlRoot = nothing
			Set xmlDoc = nothing

			Response.Write "<br><br>"
			Response.Write "<input type='button' size='35' value='Aceitar' onClick='AceitarTermoUso()'>"
		end if
	end if
end if
%>
