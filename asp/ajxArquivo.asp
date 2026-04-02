<%

Function MontaEstruturaElemento(byval xmlNode, byval codigo)
	aberto = false
	desc = trim(xmlNode.attributes.getNamedItem("DESCRICAO").value) 
	desc = Replace(desc,"""", "#QUOT#")
	
	if(xmlNode.nodeName <> "REL") then
		if (Not xmlNode.attributes.getNamedItem("CODIGO") Is Nothing) then
			tipo = "1"
			if (Not xmlNode.attributes.getNamedItem("TIPO") Is Nothing) then
				tipo = trim(xmlNode.attributes.getNamedItem("TIPO").value)
			end if
			desc = "<a class='link_classic2' href='javascript:LinkDetalhes(parent.hiddenFrame.modo_busca,1,1," & trim(xmlNode.attributes.getNamedItem("CODIGO").value) & ",1,\""link_detalhe&ve_marc=div_arquivo\""," & tipo & ");'>" & desc & "</a>"
		end if
	else
		aberto = true
	end if
	if (Not xmlNode.attributes.getNamedItem("CODIGO") Is Nothing) then
		if(trim(xmlNode.attributes.getNamedItem("CODIGO").value) = codigo) then
			desc = "<span class='destaca_palavras'>" & desc & "</span>"
			aberto = true
		end if
	end if


	cont = ""
	If xmlNode.childNodes.length > 0 Then

		cont = ""
		For i=0 to (xmlNode.childNodes.length - 1)
			cont = cont & MontaEstruturaElemento(xmlNode.childNodes(i), codigo) 
		Next
		if InStr(cont, "class='open") then
			aberto = true
		end if
		
		cont = "<ul>" & cont & "</ul>"
	End If

	classe = ""
	if(aberto) then
		classe = classe & "open "
	end if

	if (trim(xmlNode.attributes.getNamedItem("NIVEL").value) = "5") then
		classe = classe & "li-treeview-raiz "
	else
		classe = classe & "li-treeview-item "
	end if

	if(classe <> "") then
		html = "<li class='" & classe & "'>" & desc & cont & "</li>"
	else
		html = "<li>" & desc & cont & "</li>"
	end if
	

	MontaEstruturaElemento = html

End Function

sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%
cod = Request.QueryString("cod")
str_final =  "<table class='remover_bordas_padding max_height max_width justificado marc_tags_lc'>"
str_final = str_final &  "<tr>"
if isnumeric(cod) then
	'--------------------------------------------------------------------------
	'                             INICIO DA ESTRUTURA HIERARQUICA
	'--------------------------------------------------------------------------

	iIndexSrv = IntQueryString("servidor", 1)
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	
	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")
	
	xml_estrutura = ROService.MontaEstruturaHierarquica(cod, global_idioma)

	TrataErros(1)
	SET ROService = nothing
	str_tree_items = ""	
	if len(trim(xml_estrutura)) < 4  OR left(trim(uCase(xml_estrutura)),4) = "ERRO" then
		str_final = str_final &  "<tr><td><em>"&getTermo(global_idioma, 1329, "Desculpe, houve um erro na comunicação com o servidor.", 0)&"<br />"&getTermo(global_idioma, 1333, "Por favor, tente novamente mais tarde.", 0)&"</em><br/> "&xml_ficha&"</td></tr>"
	else
		str_final = str_final + "<div id='divTreeElementos' style='text-align: left; visibility: hidden'>"

		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml xml_estrutura
		Set xmlRoot = xmlDoc.documentElement
		
		If xmlRoot.childNodes.length > 0 Then

			str_final = str_final & "<ul id='tree' class='ul-treeview'>" 
			For iRaiz=0 to (xmlRoot.childNodes.length - 1)
				str_final = str_final & MontaEstruturaElemento(xmlRoot.childNodes(iRaiz), cod)
			Next
			str_final = str_final & "</ul>"

		End If

		str_final = str_final + "</div>"
	end if
else
	str_final = str_final &  "erro"
end if
serv_marc = serv_marc & "</tr>" 

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<script src="../scripts/jquery/jquery-3.2.1.min.js"></script>

<title></title>
</head>
<body>
	<div id="divConteudo">
		<script type="text/javascript">
<!--
	window.onload = function(){
		var str_final = "<%= str_final %>";

		var console = parent.document.getElementById("divArquivo");
		while(str_final.indexOf("#QUOT#") > 0) {
			str_final = str_final.replace("#QUOT#","\"");
		}

		console.innerHTML = str_final;
		
		parent.treeArquivoCarregada();
	}
	-->
</script>
	</div>
</body>
</html>