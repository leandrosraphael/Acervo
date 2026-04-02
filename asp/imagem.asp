<% Server.ScriptTimeout = 180000 %>

<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->

<%
	codItem   = Request.QueryString("item")
	codImagem = Request.QueryString("imagem")
	zoom      = Request.QueryString("zoom")
			
    sID = ROService.GetID

	Set BinImage = ROServer.CreateBinary

	if (zoom = 0) then        

		Set BinImage = ROService.GetMiniatura(CLng(codImagem))
		content      = "image/jpeg"
		sImagem      = sID & ".jpeg"

    	BinImage.SaveToFile(Server.MapPath("../temp/" & sImagem))
    	   	
    	Set objStream = Server.CreateObject("ADODB.Stream")
    	objStream.Open
    	objStream.Type = 1 
    	
    	'carrega o arquivo
    	objStream.LoadFromFile Server.MapPath("../temp/" & sImagem)
    	
    	Response.Buffer = True
    	Response.AddHeader "Content-Type", content
    	Response.Flush()
    	
    	Response.BinaryWrite objStream.Read
    	
    	objStream.Close
    	Set objStream = Nothing
    	
    	Response.Flush
    	
    	Set fs = Server.CreateObject("Scripting.FileSystemObject")
    	if fs.FileExists(Server.MapPath("../temp/" & sImagem)) then
    		fs.DeleteFile(Server.MapPath("../temp/" & sImagem))
    	end if
	else
        Set DownloadImagem = ROService.InicializaDownloadImagem(CLng(codItem), CLng(codImagem))
        QuantidadePartes = DownloadImagem.QuantidadePartes
        sGuid = DownloadImagem.Guid
        set DownloadImagem = Nothing

        if (QuantidadePartes <= 0) then

            Set BinImage = ROService.GetZoom(CLng(codItem), CLng(codImagem))

		    content      = "image/jpeg"
		    sImagem      = sID & ".jpeg"

    	    BinImage.SaveToFile(Server.MapPath("../temp/" & sImagem))
    	   	
    	    Set objStream = Server.CreateObject("ADODB.Stream")
    	    objStream.Open
    	    objStream.Type = 1 
    	
    	    'carrega o arquivo
    	    objStream.LoadFromFile Server.MapPath("../temp/" & sImagem)
    	
    	    Response.Buffer = True
    	    Response.AddHeader "Content-Type", content
    	    Response.Flush()
    	
    	    Response.BinaryWrite objStream.Read
    	
    	    objStream.Close
    	    Set objStream = Nothing
    	
    	    Response.Flush()
        else
            Set fs = Server.CreateObject("Scripting.FileSystemObject")

            ' Verifica se a pasta temp exista, caso não exista ela é criada
            If fs.FolderExists(Server.MapPath("../temp")) <> true Then
                fs.CreateFolder(Server.MapPath("../temp"))
            End if

            ' Cria a pasta com o identificador para armazenar o arquivo fragmentado
            PastaGuid = Server.MapPath("../temp") & "\" & sGuid
            fs.CreateFolder(PastaGuid)

            ' Solicita todas as partes do arquivo
            For i = 1 to QuantidadePartes
                ' Solicita o arquivo para o servidor
                Set BinParte = ROService.DownloadParteImagem(sGuid, i)

                if (BinParte.ToString <> "") then
                    BinParte.SaveToFile(PastaGuid & "\IMG" & codImagem & "._" & i)
                end if

                Set BinParte = nothing
            Next   
            
            ' Avisa o servidor que o download foi finalizado
            ' O servidor irá excluir os arquivos criados
            ROService.FinalizaDownloadImagem(sGuid)         

		    'Inicia a exibição da mídia
		    Response.Buffer = True
            Response.AddHeader "Content-Type", content
		    Response.Flush
						
            ' Envia para o browser cada uma das partes do arquivo
            For i = 1 to QuantidadePartes
			    Set objStream = Server.CreateObject("ADODB.Stream")
						    
                objStream.Open
			    objStream.Type = 1
			    objStream.LoadFromFile PastaGuid & "\IMG" & codImagem & "._" & i
						    
                Response.BinaryWrite objStream.Read
						    
                objStream.Close
			    Set objStream = Nothing

			    Response.Flush()
            Next

            'Exclui a pasta com a mídia fragmentada
            if fs.FolderExists(PastaGuid) then
                fs.DeleteFile PastaGuid & "\*.*"
                fs.DeleteFolder PastaGuid, true
            end if
        end if
	end if
	
	Set fs = nothing
	Set BinImage = nothing
	Set ROService = nothing
	Set ROServer = nothing
%>
