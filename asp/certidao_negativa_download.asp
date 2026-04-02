<% Server.ScriptTimeout = 180000 %>

<%
	sDiretorioArq = "asp" 
	nao_imprime_variaveis_globais = "1"
%>

<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->

<% 
    iCodigoCertidao = Request.QueryString("codigo")
    
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(Server.MapPath("../temp")) <> true Then
        fs.CreateFolder(Server.MapPath("../temp"))
    End if

    PastaGuid = Server.MapPath("../temp") & "/p"  

    const adSaveCreateOverWrite = 2
    const adTypeBinary = 1
    const adTypeText = 2

    Set ROService = ROServer.CreateService("Web_Consulta")                
    Set BinCertidao = ROServer.CreateBinaryType
    Set BinCertidao = ROService.BuscarCertidaoNegativa(iCodigoCertidao)
    
    if (BinCertidao.ToString <> "") then

        sArquivo = PastaGuid & iCodigoCertidao & ".pdf"        
        BinCertidao.SaveToFile(sArquivo)
                        
        Response.Buffer = False
		Response.AddHeader "Content-Type","application/octet-stream"		
        Response.AddHeader "Content-Disposition","attachment; filename=""" & iCodigoCertidao & ".pdf" & """"
		Response.Flush
						
		Set objStream = Server.CreateObject("ADODB.Stream")
						    
        objStream.Open
		objStream.Type = adTypeBinary
		objStream.LoadFromFile sArquivo
        
        Response.BinaryWrite objStream.Read
                        				    
        objStream.Close
		Set objStream = Nothing

		Response.Flush
                        
        if fs.FolderExists(PastaGuid) then
            fs.DeleteFile PastaGuid & "\*.*"                    
        end if
		
    end if

    Set ROService = nothing
    Set BinCertidao = nothing		
    Set fs = nothing
                        
%>