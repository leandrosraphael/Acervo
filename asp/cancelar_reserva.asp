<%
'-------------------------------------------------------------------------------------------------------'
'------------------------------------------- Cancela Reserva ------------------------------------------'
'-------------------------------------------------------------------------------------------------------'
codigo_reserva = Request.QueryString("codigo_reserva")
digital = Request.QueryString("digital")

if config_multi_servbib = 1 then
	iIndexSrv = Session("Servidor_Logado")

	if iIndexSrv = "" then
		iIndexSrv = 1
	end if

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%
end if

SET ROService = ROServer.CreateService("Web_Consulta")

if (digital = 1) then
	sMsg = ROService.CancelarReservaDigital(codigo_reserva,global_idioma)
else
	sMsg = ROService.CancelaReserva(codigo_reserva,global_idioma)
end if

SET ROService = nothing
TrataErros(1)

Response.Write "<script type='text/javascript'>alert('"&sMsg&"');</script>"

'Adequação Itaú
'Filtra os serviços para exibir informações somente de uma biblioteca
Response.Write "<script type='text/javascript'>"
Response.Write "	var iFiltroBib = 0;"
if ((global_numero_serie = 4090) OR (global_numero_serie = 4184)) then
	Response.Write "	if (parent.hiddenFrame.iFixarBib == 1) {"
	Response.Write "		iFiltroBib = parent.hiddenFrame.geral_bib;"
	Response.Write "	}"
end if
Response.Write "</script>"

Response.Write "<script type='text/javascript'>parent.mainFrame.location='index.asp?modo_busca='+parent.hiddenFrame.modo_busca+'&content=reservas&iBanner='+parent.hiddenFrame.iBanner + '&iIdioma='+parent.hiddenFrame.iIdioma+'&iFiltroBib='+iFiltroBib;</script>"
%>