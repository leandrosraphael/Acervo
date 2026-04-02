<%
sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<%
cod = Request.QueryString("cod")
str_final =  "<table class='remover_bordas_padding max_width max_height'>"
str_final = str_final &  "<tr>"
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
	
	xml_ficha_marc = ROService.MontaFichaMarc(cod,global_idioma,objParamDestaca)
	TrataErros(1)
	SET ROService = nothing
	
	if len(trim(xml_ficha_marc)) < 4  OR left(trim(uCase(xml_ficha_marc)),4) = "ERRO" then
		str_final = str_final &  "<tr><td><em>"&getTermo(global_idioma, 1329, "Desculpe, houve um erro na comunicação com o servidor.", 0)&"<br />"&getTermo(global_idioma, 1333, "Por favor, tente novamente mais tarde.", 0)&"</em><br/> "&xml_ficha&"</td></tr>"
	else
		str_final = str_final &  "<tr><td>"&Formata_Ficha(xml_ficha_marc,"obra",global_idioma, iIndexSrv)&"</td></tr>"
	end if
else
	str_final = str_final &  "erro"
end if
serv_marc = serv_marc & "</tr>" 

'------------------------------------------------------------
'------------------- S E R V I Ç O S ------------------------
'------------------------------------------------------------
if (global_salvar_marc = 1) then
    serv_marc = "<a class='link_serv' href=\""javascript:abreMarc("&cod&",'detalhes');\"" style='cursor:pointer'><img style='cursor:pointer;' alt='' src='imagens/icon-small-save-b.png' class='transparent-icon'>"
    serv_marc = serv_marc & "&nbsp;"&getTermo(global_idioma, 1334, "Salvar MARC", 0)&"</a>"
    str_final = str_final &  "<tr><td class='centro'>"
    str_final = str_final &  "<br/>"
    str_final = str_final &  serv_marc
    str_final = str_final &  "</td></tr>"
    str_final = str_final &  "</table><br/>"
    'Response.Write "str_final = "&str_final
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="text/javascript">
<!--
window.onload=function(){
	var str_final = "<%= str_final %>";
	var console=parent.document.getElementById("dfichaMarc");
	console.innerHTML = str_final;
}
-->
</script>
<title></title>
</head>
<body>
</body>
</html>
