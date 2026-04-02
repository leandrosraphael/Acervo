<%
'Arquivo responsável por fazer a alteração das configurações do ROServer: IP e Porta
'-----------------
'OBS IMPORTANTE 1: o valor da variável iIndexSrv deve ser definido antes do include deste arquivo
'OBS IMPORTANTE 2: o objeto Servidores é global. Sua definição é realizada no arquivo infoConn.asp
'-----------------

'Busca as informações na posição iIndexSrv do objeto Servidores
Set ServidorAtual = Servidores.ServList.Item(iIndexSrv-1)

if config_canal_conexao = "TCP" then
	ROServer.SetChannelProperty "Host", ServidorAtual.IP
	ROServer.SetChannelProperty "Port", ServidorAtual.Porta
else	
	ROServer.SetChannelProperty "TargetURL", "http://"&ServidorAtual.IP&":"&ServidorAtual.Porta&"/bin"
end if
%>
