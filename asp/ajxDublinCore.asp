<%
sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%
cod = Request.QueryString("cod")
tipo = IntQueryString("tipo", 0)

str_final =  "<table class='remover_bordas_padding max_height max_width justificado marc_tags_lc'>"

if isnumeric(cod) then
	'--------------------------------------------------------------------------
	'                             INICIO DA FICHA MARC
	'--------------------------------------------------------------------------

	iIndexSrv = IntQueryString("servidor", 1)
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
	
	On Error Resume Next
	SET ROService = ROServer.CreateService("Web_Consulta")

	'Destaque dos termos pesquisados
	redim array_campos (5)
	redim array_palavras (5)
	redim array_frase_exata (5)
		
	Set objParamDestaca = ROServer.CreateComplexType("TParamBuscaHighlight")
		
	array_campos(0) = GetPosCampoPesquisa(Request.QueryString("campo1"))
	array_campos(1) = GetPosCampoPesquisa(Request.QueryString("campo2"))
	array_campos(2) = GetPosCampoPesquisa(Request.QueryString("campo3"))
	array_campos(3) = GetPosCampoPesquisa(Request.QueryString("campo4"))
	array_campos(4) = GetPosCampoPesquisa(Request.QueryString("campo5"))
		
	sValor = RemoveUnderline(Request.QueryString("valor1"))
	array_palavras(0) = SemAspas(sValor)
	array_frase_exata(0) = EntreAspas(sValor)
	
	sValor = RemoveUnderline(Request.QueryString("valor2"))
	array_palavras(1) = SemAspas(sValor)
	array_frase_exata(1) = EntreAspas(sValor)
	
	sValor = RemoveUnderline(Request.QueryString("valor3"))
	array_palavras(2) = SemAspas(sValor)
	array_frase_exata(2) = EntreAspas(sValor)
	
	sValor = RemoveUnderline(Request.QueryString("valor4"))
	array_palavras(3) = SemAspas(sValor)
	array_frase_exata(3) = EntreAspas(sValor)
	
	sValor = RemoveUnderline(Request.QueryString("valor5"))
	array_palavras(4) = SemAspas(sValor)
	array_frase_exata(4) = EntreAspas(sValor)
			
	'CRIANDO ARRAYS RO
	Set aiCampos = ROServer.CreateComplexType("TInteiro")
	aiCampos.Param0 = array_campos(0)
	aiCampos.Param1 = array_campos(1)
	aiCampos.Param2 = array_campos(2)
	aiCampos.Param3 = array_campos(3)
	aiCampos.Param4 = array_campos(4)
	
	Set aiPalavras = ROServer.CreateComplexType("TString")
	aiPalavras.Param0 = array_palavras(0)
	aiPalavras.Param1 = array_palavras(1)
	aiPalavras.Param2 = array_palavras(2)
	aiPalavras.Param3 = array_palavras(3)
	aiPalavras.Param4 = array_palavras(4)

	Set aiFrase = ROServer.CreateComplexType("TString")
	aiFrase.Param0 = array_frase_exata(0)
	aiFrase.Param1 = array_frase_exata(1)
	aiFrase.Param2 = array_frase_exata(2)
	aiFrase.Param3 = array_frase_exata(3)
	aiFrase.Param4 = array_frase_exata(4)

	objParamDestaca.asPalavras   = aiPalavras
	objParamDestaca.asFraseExata = aiFrase
	objParamDestaca.aiCampos     = aiCampos
	'Fim do destaque os termos pesquisados
	
	xml_dublin_core = ROService.MontarDublinCore(cod, objParamDestaca)

	TrataErros(1)

	Set objParamDestaca = nothing
	SET ROService = nothing
		
	if len(trim(xml_dublin_core)) < 4  then
		str_final = str_final &  "<tr><td><em>"&getTermo(global_idioma, 1329, "Desculpe, houve um erro na comunicação com o servidor.", 0)&"<br />"&getTermo(global_idioma, 1333, "Por favor, tente novamente mais tarde.", 0)&"</em><br/></td></tr>"
	else
		sDC = "<table style='width: 100%'>"

		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml xml_dublin_core
		Set xmlRoot = xmlDoc.documentElement

		For Each xmlTermo In xmlRoot.childNodes
			sDC = sDC & "<tr>"
			sDC = sDC & "<td style='width: 100px; font-weight: bold;'>" & xmlTermo.nodeName & "</td>"
			sDC = sDC & "<td>" 
			
			if ((xmlTermo.nodeName = "relation") or (xmlTermo.nodeName = "identifier")) and (InStr(xmlTermo.text, "http") > 0) then
				valor = replace(replace(xmlTermo.text, "<", "&lt;"), ">", "&gt;")
				sDC = sDC & "<a class='link_classic2' href='" & valor & "' target='_blank'>" & valor & "</a>"
			else
				sDC = sDC & TrocaTagMarcador(replace(replace(xmlTermo.text, "<", "&lt;"), ">", "&gt;"))
            end if
            sDC = sDC & "</td>"
			sDC = sDC & "</tr>"
		Next

		sDC = sDC & "</table>"

		Set xmlRoot = nothing
		Set xmlDoc = nothing
        sDC =  replace(replace(replace(sDC,chr(10),"<br />"),"\","\\"),"""","\""")
		str_final = str_final &  "<tr><td class='td-marc-tags'>"&sDC&"</td></tr>"
	end if
else
	str_final = str_final &  "erro"
end if

str_final = str_final & "</table>"

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<script type="text/javascript">
<!--
window.onload=function(){
	var str_final = "<%= str_final %>";
	var console = parent.document.getElementById("dDublinCore");
	console.innerHTML = str_final;
}
-->
</script>
<title></title>
</head>
<body>
</body>
</html>