<!-- #include file="servidores.asp" -->
<%

'Define como buscar o arquivo XML com a lista de servidores
sLibDir = ""
select case sDiretorioArq 
	case "root"
		sLibDir = "libasp\"
	case "asp"
		sLibDir = "..\libasp\"
	case "externo"
		sLibDir = "..\libasp\"
	case "idiomas"
		sLibDir = "..\libasp\"
	case "orcamento"
		sLibDir = "..\libasp\"
end select 

'Cria lista de servidores
Set Servidores = new ServidoresClass

if config_multi_servbib = 1 then		
	'Define como buscar o arquivo XML com a lista de servidores
	sXMLDir = sLibDir & "Servidores.xml"
	
	Set xmlSrv = CreateObject("Microsoft.xmldom")
	xmlSrv.async = False
	xmlSrv.load(Server.MapPath(sXMLDir))
	
	Set xmlRoot = xmlSrv.documentElement
		
	if xmlRoot.nodeName = "servidores" then	
			
		iNumGrupo = xmlRoot.attributes.getNamedItem("qtde").Value
		sNomeGrupo = xmlRoot.attributes.getNamedItem("nome").Value
		
		Servidores.Nome = sNomeGrupo
					
		'Percorre a lista de servidores
		For Each xmlServidor In xmlRoot.childNodes
			if xmlServidor.nodeName = "servidor" then
				'Cria um servidor		
				Set Servidor = new ServidorClass			
				'Percorre os campos do servidor
				For Each xmlCamposServidor In xmlServidor.childNodes				
					if xmlCamposServidor.nodeName = "ip" then
						Servidor.IP = xmlCamposServidor.text						
					end if
					if xmlCamposServidor.nodeName = "porta" then
						Servidor.Porta = xmlCamposServidor.text
					end if
					if xmlCamposServidor.nodeName = "nome" then
						Servidor.Nome = xmlCamposServidor.text
					end if
					if xmlCamposServidor.nodeName = "bibliotecas" then
						'Percorre a lista de bibliotecas
						For Each xmlBiblioteca In xmlCamposServidor.childNodes
							if xmlBiblioteca.nodeName = "biblioteca" then
								For Each xmlCampoBiblioteca In xmlBiblioteca.childNodes
									if xmlCampoBiblioteca.nodeName = "codigo" then
										CodBib = xmlCampoBiblioteca.text
									end if
									if xmlCampoBiblioteca.nodeName = "nome" then
										NomeBib = xmlCampoBiblioteca.text
									end if
								Next
								'Adiciona uma biblioteca na lista
								Servidor.AddBiblioteca CodBib, NomeBib
							end if
						Next
					end if
				Next
				'Adiciona o servidor na lista
				Servidores.AddServidor Servidor
			end if
		Next
	end if
else
	'Cria um servidor
	Set Servidor = new ServidorClass

	if ((habilita_espelho_servbib = 0) or ((habilita_espelho_servbib = 1) and (Session("espelho") <> "1"))) then
		'IP e Porta do servibib principal
		Servidor.IP = config_servbib_ip
		Servidor.Porta = config_servbib_porta
	else
		'IP e Porta do servbib espelho
		Servidor.IP = config_servbibesp_ip
		Servidor.Porta = config_servbibesp_porta
	end if	
	
	'Adiciona o servidor na lista
	Servidores.AddServidor Servidor		
end if
%>

