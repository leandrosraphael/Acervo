<%
'***********************************************************************************************
' Build do Terminal Web: 2009.3.xx.yy
' xx: Build mínimo do SrvAcervo necessário para o Terminal Web
' yy: Build do Terminal Web
'***********************************************************************************************
build_web = "2009.3.108.34"
'***********************************************************************************************
' Customização do Layout
'***********************************************************************************************
config_nome_cliente = "Museu da Justiça – Centro Cultural do Poder Judiciário"
config_customizacao_layout = 0
'***********************************************************************************************
' Configurações de Acesso
'***********************************************************************************************
'IP do servidor de aplicação
config_ip = "tjerj2119AVM"
'Porta de comunicação com o servidor de aplicação          
config_porta = "8098"
'Canal de comunicação (TROWinInetHTTPChannel/HTTP/TCP)
config_canal_conexao = "TROWinInetHTTPChannel"   
'Uso de mídias offline: todas as mídias devem ser copiadas para a pasta temp
config_midia_offline = 0
%>
