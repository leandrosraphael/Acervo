<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
<% sDiretorioArq="asp" %>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<title>::<%=getTermo(global_idioma, 963, "Minha seleção", 0)%></title>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link rel="stylesheet" href="../inc/estilo_imp.css" />
<link rel="stylesheet" href="../inc/contraste.css" />

</head>
<body onload="parent.parent.fechaLoadingPopup();">
<table class="tabela_imp">
<tr>
<td class="td_imp td_right_top">
<%
codigoFavorito = IntQueryString("codigoFavorito", 0) 
iIndexSrv = IntQueryString("Servidor", 1)
'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
%><!-- #include file="../libasp/updChannelProperty.asp" --><%

if(config_multi_servbib = 1)then
	sNomeBib = arNomeSrv(iIndexSrv)
else
	sNomeBib = global_instituicao
end if

sSessao = "mysel"&iIndexSrv
mysel = Session(sSessao)

if (codigoFavorito > 0) then
	Response.write "<table class='table_resultados' style='border-spacing: 1px; padding: 4px; border: 0'>"
	Response.write "<tr>"
	Response.write "	<td class='td_resultados td_center_top'>"
	Response.write "		<br /><div class='centro' style='padding-bottom: 10px;border-bottom: 1px solid #000;margin-bottom: 10px;'><br /><b>"&sNomeBib&"</b></div>"

	'*******************************************************
	Set ROService = ROServer.CreateService("Web_Consulta")
	sXMLFichas = ROService.MontaFichaFavoritos(codigoFavorito, global_idioma)
	Set ROService = nothing
	'*******************************************************
	  
	if len(trim(sXMLFichas)) = 0 then
		Response.write "	<p class='centro'>"&getTermo(global_idioma, 1475, "Nenhuma seleção no momento", 0)&"</p>"
	else 

		Response.Write "	<table style='width: 100%; border-spacing: 1px; padding: 4px; background-color: #f2f2f2'>"
		Response.Write "	<tr style='height: 26px'>"
		Response.Write "		<td class='esquerda' style='width: 100%'>&nbsp;&nbsp;"
		
		if (left(sXMLFichas,5) = "<?xml") then
			Set xmlDoc = CreateObject("Microsoft.xmldom")
			xmlDoc.async = False
			xmlDoc.loadxml sXMLFichas
			Set xmlRoot = xmlDoc.documentElement
			if xmlRoot.nodeName = "Pagina" then
				quantidade = xmlRoot.attributes.getNamedItem("QUANTIDADE").value
			end if
			Set xmlDoc = nothing
			Set xmlRoot = nothing		
		end if	
	
		if quantidade = 1 then
			msg_sel = getTermo(global_idioma, 2945, "%s material selecionado", 0)
			msg_sel = Format(msg_sel, "<b>"&quantidade&"</b>")
			Response.Write msg_sel
		else
			msg_sel = getTermo(global_idioma, 2946, "%s materiais selecionados", 0)
			msg_sel = Format(msg_sel, "<b>"&quantidade&"</b>")
			Response.Write msg_sel
		end if

		Response.Write "		&nbsp;&nbsp;&nbsp;&nbsp;"&now()&"</td>"
		Response.Write "</tr></table>"
	
		%><!-- #include file="grid_ficha_sels_imp.asp" --> <%
	
	end if

	Response.write "</tr>"
	Response.write "</table>"
	Response.write "</td>"
	Response.write "</tr>"
	Response.write "<tr class='td_right_bottom'><td>"
	Response.write "SophiA Biblioteca"

elseif len(trim(mysel)) > 0 then
	
	Response.write "<table class='table_resultados' style='border-spacing: 1px; padding: 4px; border: 0'>"
	Response.write "<tr>"
	Response.write "	<td class='td_resultados td_center_top'>"
	Response.write "		<br /><div class='centro' style='padding-bottom: 10px;border-bottom: 1px solid #000;margin-bottom: 10px;'><br /><b>"&sNomeBib&"</b></div>"

	arMySel = split(mysel,",")
	sPagina = ""
	
	for i = lbound(arMySel) to ubound(arMySel)
		arTemp   = split(arMySel(i),".")
		sPagina  = sPagina & arTemp(0) & ","
		Sel_Tipo = arTemp(1)
		
		if(cInt(Sel_Tipo) = 2) then
			iNumLeg = iNumLeg + 1
		end if
	next

	'*******************************************************
	Set ROService = ROServer.CreateService("Web_Consulta")
	sXMLFichas = ROService.MontaFichaMinhaSelecao(sPagina, global_idioma)
	Set ROService = nothing
	'*******************************************************
	  
	if len(trim(sXMLFichas)) = 0 then
		Response.write "	<p class='centro'>"&getTermo(global_idioma, 1475, "Nenhuma seleção no momento", 0)&"</p>"
	else 
		arMySel = split(mysel,",")
		iNumSel = ubound(arMySel) + 1

		Response.Write "	<table style='width: 100%; border-spacing: 1px; padding: 4px; background-color: #f2f2f2'>"
		Response.Write "	<tr style='height: 26px'>"
		Response.Write "		<td class='esquerda' style='width: 100%'>&nbsp;&nbsp;"
		
		if iNumSel = 1 then
			msg_sel = getTermo(global_idioma, 2945, "%s material selecionado", 0)
			msg_sel = Format(msg_sel, "<b>"&iNumSel&"</b>")
			Response.Write msg_sel
		else
			msg_sel = getTermo(global_idioma, 2946, "%s materiais selecionados", 0)
			msg_sel = Format(msg_sel, "<b>"&iNumSel&"</b>")
			Response.Write msg_sel
		end if

		Response.Write "		&nbsp;&nbsp;&nbsp;&nbsp;"&now()&"</td>"
		Response.Write "</tr></table>"
	
		%><!-- #include file="grid_ficha_sels_imp.asp" --> <%
	
	end if

	Response.write "</tr>"
	Response.write "</table>"
	Response.write "</td>"
	Response.write "</tr>"
	Response.write "<tr class='td_right_bottom'><td>"
	Response.write "SophiA Biblioteca"
else
	Response.write "<p class='centro'>"&getTermo(global_idioma, 1475, "Nenhuma seleção no momento", 0)&"</p>"
end if
%>
</td>
</tr>
</table>
</body>
</html>