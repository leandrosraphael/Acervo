<%
    nao_imprime_variaveis_globais = "1"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<!-- #include file="../libasp/roclient.asp" -->

<%

codigoMidia = request.QueryString("codigoMidia")

QuantidadePaginas = ROService.InicializarMidiaPaginacaoPDF(CLng(codigoMidia))

numeroPaginas = Clng(QuantidadePaginas)

Set QuantidadePaginas = nothing
Set ROService = nothing
Set ROServer  = nothing

if (numeroPaginas <= 0) then
    Pagina = 0
else
    Pagina = 1
end if
 
response.Write numeroPaginas
%>