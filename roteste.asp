<!-- #include file="config.asp" -->
<%	
	On Error Resume Next
	set ROServer = Server.CreateObject("RemObjects.SDK.COMServer")
	
	if err.number <> 0 then
		response.write err.source&" - "&err.description
		Response.End()
	end if
	
	ROServer.MessageType = "BinMessage"
	ROServer.ChannelType = config_canal_conexao
	
	ROServer.SetChannelProperty "TargetURL", "http://"&config_ip&":"&config_porta&"/bin"
	
	Set ROService = ROServer.CreateService("srvWeb")	

	if err.number <> 0 then
		Response.Write err.source&" - "&err.description
		Response.End()
	end if
	
	numeroSerie = ROService.GetNumeroDTA
	
	if err.number <> 0 then
		Response.Write err.source&" - "&err.description
		Response.End()
	end if
	
	Response.Write " NS: " & numeroSerie
	
	Set ROService = nothing
	Set ROServer = nothing
%>