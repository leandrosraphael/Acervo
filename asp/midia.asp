<% Server.ScriptTimeout = 180000 %>

<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->

<%
	codMidia = Request.QueryString("codigo")
	sExt     = Request.QueryString("ext")
	sArq     = Request.QueryString("arq")
	sID = CStr(codMidia) & "." & sExt
    usuario = Session("codigo")

	'******************************************
	'Contagem de acesso
	'******************************************
	ROService.ContaAcessosMidias CLng(codMidia), CLng(usuario)

    Set DownloadImagem = ROService.InicializaDownloadMidia(CLng(codMidia))
    QuantidadePartes = DownloadImagem.QuantidadePartes
    sGuid = DownloadImagem.Guid
    set DownloadImagem = Nothing
    
    
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
             BinParte.SaveToFile(PastaGuid & "\IMG" & codMidia & "._" & i)
         end if

         Set BinParte = nothing
     Next   
     
     ' Avisa o servidor que o download foi finalizado
     ' O servidor irá excluir os arquivos criados
     ROService.FinalizaDownloadImagem(sGuid)         
    
     'Inicia a exibição da mídia
     Response.Buffer = True
     Response.AddHeader "Content-Type","application/octet-stream"
     Response.AddHeader "Content-Disposition","attachment; filename=""" & sArq & "." & sExt & """"
     Response.Flush
	   						
     ' Envia para o browser cada uma das partes do arquivo
     For i = 1 to QuantidadePartes
	    Set objStream = Server.CreateObject("ADODB.Stream")
				    
         objStream.Open	 
	     objStream.Type = 1
	     objStream.LoadFromFile PastaGuid & "\IMG" & codMidia & "._" & i 
		
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


%>
