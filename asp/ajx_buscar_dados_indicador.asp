<%
	sDiretorioArq = "asp"
    nao_imprime_variaveis_globais = "1"
%>

<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<%
iIndexSrv = 1
%>

<!-- #include file="../libasp/updChannelProperty.asp" -->

<%

indicador  = Request.QueryString("indicador")

On Error Resume Next
SET ROService = ROServer.CreateService("Web_Consulta")

if (indicador = "programa") then
	dadosGraficoJson = ROService.ObterDadosGraficoTeseDissertacaoPrograma()
elseif (indicador = "material") then
	dadosGraficoJson = ROService.ObterDadosGraficoTeseDissertacaoMaterial()
elseif (indicador = "curso") then
	dadosGraficoJson = ROService.ObterDadosGraficoTrabalhoGraduacaoCurso()
elseif (indicador = "empresa") then
	dadosGraficoJson = ROService.ObterDadosGradeAcessoEmpresa()
elseif (indicador = "geral_material") then
    dadosGraficoJson = ROService.ObterDadosGraficoMaterialAno()
elseif (indicador = "reset") then
	ROService.ResetarConfiguracaoCacheIndicadorWeb()
else
	dadosGraficoJson = ""
end if

SET ROService = nothing

Response.ContentType = "application/json"
Response.Write(dadosGraficoJson)

%>