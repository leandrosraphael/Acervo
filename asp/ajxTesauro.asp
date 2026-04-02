<%
Function MontaEstruturaThesauro(byval xmlNode)
	strResult = ""
	strLink = ""
	if (xmlNode.hasChildNodes()) then
		codigo = xmlNode.attributes.getNamedItem("CODIGO").value
		termo = Trim(xmlNode.attributes.getNamedItem("TERMO").value)
		grupo = Trim(xmlNode.attributes.getNamedItem("GRUPO").value)
		strLink = "<a class='link_classic2' href=\""javascript:LinkBuscaAssunto(parent.hiddenFrame.modo_busca,"&codigo&",'"&termo&"',"&iIndexSrv&");\"" style='cursor:pointer;' title=\"""&getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"...\"">"&grupo& " " &termo&"</a>"
		strResult = strResult & "<li class='li-treeview-item'>"&strLink&"<ul>"
		for each xmlFilho in xmlNode.childNodes
			strResult = strResult & Trim(MontaEstruturaThesauro(xmlFilho))
		next
		strResult = strResult & "</ul>"
	else
		codigo = xmlNode.attributes.getNamedItem("CODIGO").value
		if (codigo <> "") then
			termo = Trim(xmlNode.attributes.getNamedItem("TERMO").value)
			grupo = Trim(xmlNode.attributes.getNamedItem("GRUPO").value)
			strLink = "<a class='link_classic2' href=\""javascript:LinkBuscaAssunto(parent.hiddenFrame.modo_busca,"&codigo&",'"&termo&"',"&iIndexSrv&");\"" style='cursor:pointer;' title=\"""&getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"...\"">"&grupo& " " &termo&"</a>"
			strResult = strResult & "<li class='li-treeview-item'>"&strLink
			strResult = strResult&"</li>"
		end if
	end if

	MontaEstruturaThesauro = strResult
End Function

sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%

cod = Request.QueryString("cod")
tipo = Request.QueryString("tipo")

iIndexSrv = IntQueryString("Servidor",1)
strTreeThesauro = "<table class='remover_bordas_padding max_height max_width justificado tab-tesauro'>"
if isnumeric(cod) then
	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	xml_thesaurus = ROService.MontarThesaurusAutoridade(cod,global_idioma)
	SET ROService = nothing
	TrataErros(1)
	qtdRelacionada = 0
	qtdEstrutura = 0
	if len(trim(xml_thesaurus)) < 4  OR left(trim(uCase(xml_thesaurus)),4) = "ERRO" then
		strTreeThesauro = strTreeThesauro & "<tr><td><em>"&getTermo(global_idioma, 1329, "Desculpe, houve um erro na comunicação com o servidor.", 0)&"<br />"&getTermo(global_idioma, 1333, "Por favor, tente novamente mais tarde.", 0)&"</em><br /> "&xml_thesaurus&"</td></tr>"
	else
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml xml_thesaurus
		strTreeThesauro = strTreeThesauro & "<div id='div-thesauro-Elementos' style='text-align: left; visibility: hidden'>"
		
		if (Tipo = "0") then
		   sTituloTesauro = getTermo(global_idioma, 460, "Estrutura do termo", 0)
		else
		   sTituloTesauro = getTermo(global_idioma, 462, "Estruturas relacionadas ao termo", 0)
		end if
		sTermoTesauro = ""

		Set xmlRoot = xmlDoc.documentElement
		if (xmlRoot.NodeName = "THESAURO_SOPHIA") then
			sTermoTesauro = xmlRoot.attributes.getNamedItem("DESC_AUTORIDADE").value
			if (Tipo = "1") then
			   strTreeThesauro = strTreeThesauro & "<tr><td colspan='4' class='td_center_top td_marc' style='background-color: #fff;'><h4 style='text-align: left;'>"&sTermoTesauro&"</h4>"
			else
			   strTreeThesauro = strTreeThesauro & "<tr><td colspan='4' class='td_center_top td_marc' style='background-color: #fff;'>"
			end if
	
			qtdRelacionada = xmlRoot.attributes.getNamedItem("QTD_RELACIONADA").value
			qtdEstrutura = xmlRoot.attributes.getNamedItem("QTD_ESTRUTURA").value
			tesauro_relacionado = xmlRoot.attributes.getNamedItem("THESAURO_RELACIONADO").value

			strTreeThesauro = strTreeThesauro & "<div id='div-thesauro-detalhe'><ul id='tree-thesauro'>"
			For Each xmlTermoTopico In xmlRoot.childNodes
				if (xmlTermoTopico.NodeName = "TERMOS_TOPICOS") then
					codigo = xmlTermoTopico.attributes.getNamedItem("CODIGO").value
					if (codigo <> "") then
						termo = Trim(xmlTermoTopico.attributes.getNamedItem("TERMO").value)
						grupo =  Trim(xmlTermoTopico.attributes.getNamedItem("GRUPO").value)
						strLink = "<a class='link_classic2' href=\""javascript:LinkBuscaAssunto(parent.hiddenFrame.modo_busca,"&codigo&",'"&termo&"',"&iIndexSrv&");\"" style='cursor:pointer;' title=\"""&getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"...\"">"&grupo&" "&termo&"</a>"
						strTreeThesauro = strTreeThesauro & "<li class='li-treeview-item' >"&strLink&"</li>"
					end if
				end if
			next

			For Each xmlNos In xmlRoot.childNodes

				if (xmlNos.NodeName = "TERMOS_SOPHIA") then
					For Each xmlTermosSophia In xmlNos.childNodes
						codigo = xmlTermosSophia.attributes.getNamedItem("CODIGO").value
						if (codigo <> "") then
							termo = Trim(xmlTermosSophia.attributes.getNamedItem("TERMO").value)
							grupo = Trim(xmlTermosSophia.attributes.getNamedItem("GRUPO").value)
							strLink = "<a class='link_classic2' href=\""javascript:LinkBuscaAssunto(parent.hiddenFrame.modo_busca,"&codigo&",'"&termo&"',"&iIndexSrv&");\"" style='cursor:pointer;' title=\"""&getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"...\"">"&grupo& " " &termo&"</a>"
							strTreeThesauro = strTreeThesauro & "<li class='li-treeview-raiz'>" & strLink
							if (xmlTermosSophia.hasChildNodes()) then
								strTreeThesauro = strTreeThesauro & "<ul>"
								for each xmlFilho in xmlTermosSophia.childNodes
									strTreeThesauro = strTreeThesauro & MontaEstruturaThesauro(xmlFilho)
								next
								strTreeThesauro = strTreeThesauro & "</ul>"
							end if
							strTreeThesauro = strTreeThesauro & "</li>"
						end if
					next
				end if

			next

			For Each xmlTermoRelacionado In xmlRoot.childNodes
				if (xmlTermoRelacionado.NodeName = "TERMOS_RELACIONADOS") then
					codigo = xmlTermoRelacionado.attributes.getNamedItem("CODIGO").value
					if (codigo <> "") then
						termo = Trim(xmlTermoRelacionado.attributes.getNamedItem("TERMO").value)
						grupo =  Trim(xmlTermoRelacionado.attributes.getNamedItem("GRUPO").value)
						strLink = "<a class='link_classic2' href=\""javascript:LinkBuscaAssunto(parent.hiddenFrame.modo_busca,"&codigo&",'"&termo&"',"&iIndexSrv&");\"" style='cursor:pointer;' title=\"""&getTermo(global_idioma, 1563, "Buscar todos os registros deste assunto", 0)&"...\"">"&grupo&" "&termo&"</a>"
						strTreeThesauro = strTreeThesauro & "<li class='li-treeview-item' >"&strLink&"</li>"
					end if
				end if
			next
			strTreeThesauro = strTreeThesauro + "</ul></div>"
		end if
		response.write "</div>"
	end if
	strTreeThesauro = strTreeThesauro + "</td></tr>"
else
	strTreeThesauro = strTreeThesauro +  "<tr><td>ERRO</td></tr>"
end if

if (qtdRelacionada = 0) and (qtdEstrutura = 0) then
	strTreeThesauro = strTreeThesauro & "<tr><td colspan='4' class='td_left_middle'><b>"&getTermo(global_idioma, 458, "Não há estrutura cadastrada nem relacionada para esse termo.", 0) & "</b></td></tr><br />"
	strTreeThesauro = strTreeThesauro & "<tr><td colspan='4' class='td_left_middle'>&nbsp;</b></td></tr>"
	strTreeThesauro = strTreeThesauro & "<tr><td colspan='4' class='td_left_middle'><div class='background_aba_ativa centro'>"
	strTreeThesauro = strTreeThesauro & "</div></td></tr>"
else
	strTreeThesauro = strTreeThesauro & "<tr><td colspan='4' class='td_left_middle'><b>"&getTermo(global_idioma, 461, "Estruturas relacionadas", 0) & ": " & qtdRelacionada & "</b></td></tr><br />" 
	strTreeThesauro = strTreeThesauro & "<tr><td colspan='4' class='td_left_middle'><b>"&getTermo(global_idioma, 459, "Itens da estrutura selecionada", 0) & ": " & qtdEstrutura & "</b></td></tr>"
	strTreeThesauro = strTreeThesauro & "<tr><td colspan='2' class='td_left_middle'><div class='background_aba_ativa centro'>"
	strTreeThesauro = strTreeThesauro & "<span title='"&getTermo(global_idioma, 472, "Passar para a estrutura anterior", 0)&"' class='cursor_padrao transparent-icon span_imagem icon_16 icon-small-previous' onclick='javascript:ThesauroAnterior();'>&nbsp;&nbsp;</span>"
	strTreeThesauro = strTreeThesauro & "<span title='"&getTermo(global_idioma, 473, "Passar para a próxima estrutura", 0)&"' class='cursor_padrao transparent-icon span_imagem icon_16 icon-small-next' onclick='javascript:ThesauroProximo();'>&nbsp;&nbsp;</span>"
	strTreeThesauro = strTreeThesauro & "</div></td></tr>"
end if

strTreeThesauro = strTreeThesauro & "</table></div><br />"

%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
		<meta http-equiv="X-UA-Compatible" content="IE=edge" />
		<title></title>
	</head>
	
	<body>
		<div id="divConteudo">
			<script src="../scripts/jquery/jquery-3.2.1.min.js"></script>
			<script src="../scripts/jquery.treeview.js"></script>
			<script type="text/javascript">

				window.onload = function () {
					var str_final = "<%= strTreeThesauro %>";
					var TituloTesauro = "<%= sTituloTesauro %>";
					var TermoTesauro = "<%= sTermoTesauro %>";
					var TesauroRelacionado = "<%= tesauro_relacionado %>";
					var arvoreThesaurus = parent.document.getElementById("div-thesauro-autoridade");
					var codigoTesauro = "<%= cod %>";
					var TipoBusca = "<%= tipo %>";

					if (parent.hiddenFrame != null) {
						var Frame = parent.hiddenFrame;
					} else if (parent.parent.hiddenFrame != null) {
						var Frame = parent.parent.hiddenFrame;
					} else {
						var Frame = parent.parent.parent.hiddenFrame;
					}

					arvoreThesaurus.innerHTML = str_final;

					parent.treeThesauroCarregado();

					if (TipoBusca == "0") {
					   if (TesauroRelacionado != "") {
							Frame.other_thesauro = TesauroRelacionado.split(",");
					   } else {
						   Frame.other_thesauro = [];
					   }
					   Frame.other_thesauro.push(Frame.main_thesauro);
					   Frame.other_thesauro_atual = Frame.other_thesauro.length - 1;
					}
					
					parent.document.getElementById("titulo-tesauro").innerHTML = TituloTesauro;
					if (codigoTesauro == Frame.main_thesauro) {
						parent.document.getElementById("termo-tesauro").innerHTML = ' "' + TermoTesauro + '"';
					}
					
				}
	
			</script>
		</div>
	</body>
</html>