<%
	sMsgErro = ""

	On Error Resume Next
	Set ROServer = Server.CreateObject("RemObjects.SDK.COMServer")
	
	If err.number <> 0 then
		sMsgErro = TrataErros
	else
		ROServer.MessageType = "BinMessage"
		ROServer.ChannelType = config_canal_conexao
		
		if config_canal_conexao = "TCP" then
			ROServer.SetChannelProperty "Host", config_ip
			ROServer.SetChannelProperty "Port", config_porta
		else	
			ROServer.SetChannelProperty "TargetURL", "http://"&config_ip&":"&config_porta&"/bin"
		end if
	
		Set ROService = ROServer.CreateService("srvWeb")
		
		If err.number <> 0 then
			sMsgErro = TrataErros
		else
			ROService.Hello
			
			If err.number <> 0 then
				sMsgErro = TrataErros
			end if
		end if
	end if
%>