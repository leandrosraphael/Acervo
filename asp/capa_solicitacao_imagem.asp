<% Server.ScriptTimeout = 180000 %>

<% 
	sDiretorioArq = "asp" 
	nao_imprime_variaveis_globais = "1"
%>

<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<% 
	iIndexSrv = IntQueryString("servidor", 1)
%><!-- #include file="../libasp/updChannelProperty.asp" --><%	

    codMidia  = Request.QueryString("codigo")
	sArquivo = ""
	
	On Error Resume Next
	Set ROService = ROServer.CreateService("Web_Consulta")	
	Set BinFile   = ROServer.CreateBinaryType
	Set BinFile   = ROService.CapaSolicitacaoImagem(CLng(codMidia))
	
	if (BinFile.ToString = "") then
		'Servidor não encontrou a capa
		Response.Buffer = True
		Response.AddHeader "Content-Type","image/jpeg"
		Response.Flush
		
		img_capa = "../imagens/capa_fantasia.png"
		
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open
		objStream.Type = 1
		objStream.LoadFromFile Server.MapPath(img_capa)
		Response.BinaryWrite objStream.Read
		objStream.Close
		Set objStream = Nothing
		Response.Flush	
			
		Set BinFile   = nothing
		Set ROService = nothing
		Set ROServer  = nothing
	else				
		sID = "M"&codMidia&".jpg"
		
		BinFile.SaveToFile(Server.MapPath(global_midia_caminho_fisico & "/" & sID))
	
		if (err.number <> 0) then
			'Ocorreu um erro ao gravar a capa
			Response.Buffer = True
			Response.AddHeader "Content-Type","image/jpeg"
			Response.Flush
			
			img_capa = "../imagens/capa_fantasia.png"
			
			Set objStream = Server.CreateObject("ADODB.Stream")
			objStream.Open
			objStream.Type = 1
			objStream.LoadFromFile Server.MapPath(img_capa)
			Response.BinaryWrite objStream.Read
			objStream.Close
			Set objStream = Nothing
			Response.Flush	
		else
			'Inicia a exibição da capa
			Response.Buffer = True
			Response.AddHeader "Content-Type","image/jpeg"
			Response.Flush
			
			Set objStream = Server.CreateObject("ADODB.Stream")
			objStream.Open
			objStream.Type = 1
			objStream.LoadFromFile Server.MapPath(global_midia_caminho_fisico & "/" & sID)
			Response.BinaryWrite objStream.Read
			objStream.Close
			Set objStream = Nothing
			Response.Flush	
			
			'Exclui a mídia
			Set fs = Server.CreateObject("Scripting.FileSystemObject")
			if fs.FileExists(Server.MapPath(global_midia_caminho_fisico & "/" & sID)) then
		 	   fs.DeleteFile(Server.MapPath(global_midia_caminho_fisico & "/" & sID))
			end if
			Set fs = nothing
		end if
	
		Set BinFile   = nothing
		Set ROService = nothing
		Set ROServer  = nothing
	end if
%>