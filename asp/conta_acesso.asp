<% Server.ScriptTimeout = 180000 %>

<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->

<%
	codigoMidia = Request.QueryString("midia")
    codigoUsuario = Session("codigo")

    Set ROService = ROServer.CreateService("Web_Consulta")
	ROService.ContaAcessosMidias CLng(codigoMidia), CLng(codigoUsuario)
	Set ROService = nothing
%>
