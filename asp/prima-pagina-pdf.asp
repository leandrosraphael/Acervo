<%
	'Desabilitando cache pois estava influenciando na passagem de parâmetros no site
	'Ao recarregar a mesma página os dados nem sempre eram atualizados
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	
    nao_imprime_variaveis_globais = "1"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<!-- #include file="../libasp/roclient.asp" -->


<%
if (request.QueryString("codigoMidia") <> "") then
    Pagina = request.QueryString("pagina")
    CodigoMidia = request.QueryString("codigoMidia")
end if

if (CodigoMidia <> "") and (Pagina > 0) then
    
    err.Clear
    On Error Resume Next    
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    
    If fs.FolderExists(Server.MapPath("../temp")) <> true Then
        fs.CreateFolder(Server.MapPath("../temp"))
    End if

    sPagina = ""
    lim = 6-Len(CStr(Pagina))
	For i = 1 to lim
		sPagina = sPagina + "0"
	Next
    sPagina = sPagina + cStr(Pagina)

    sArquivo = Server.MapPath("../temp") & "/p" &CodigoMidia & sPagina & ".png"

	Set Imagem = ROServer.CreateBinaryType
    Set Imagem = ROService.ObterImagemMidiaPDF(Clng(CodigoMidia), Clng(Pagina))

    const adSaveCreateOverWrite = 2
    const adTypeBinary = 1
    const adTypeText = 2
    'Inicia a exibição da imagem
    if (Imagem.ToString = "") then
			
	    Response.Buffer = True
	    Response.AddHeader "Content-Type","image/jpeg"
	    Response.Flush
		
        Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open
        objStream.Type = adTypeText			    
        objStream.WriteText ""
        Response.BinaryWrite objStream.Read

		objStream.Close
        Set objStream = Nothing
		
        Response.Flush		
    else
        Imagem.SaveToFile(sArquivo)
    	
		Response.Buffer = False
		Response.AddHeader "Content-Type","application/octet-stream"
		Response.AddHeader "Content-Disposition","attachment; filename=""" & sArquivo & """"
		Response.Flush
						
		Set objStream = Server.CreateObject("ADODB.Stream")
						    
        objStream.Open
		objStream.Type = adTypeBinary
		objStream.LoadFromFile sArquivo
						    
        if (objStream.Size > 1048576) then
            iTamanhoImagem = 1048576 '1MB
            For i = 1 To objStream.Size \ iTamanhoImagem
                Response.BinaryWrite objStream.Read(iTamanhoImagem)
            Next
            if (objStream.Size Mod iTamanhoImagem) <> 0 then
                Response.BinaryWrite objStream.Read(objStream.Size Mod iTamanhoImagem)
            end If
        else
            Response.BinaryWrite objStream.Read
        end if
						    
        objStream.Close
		Set objStream = Nothing

		Response.Flush
        
        fs.DeleteFile sArquivo
	end if
		
    Set Imagem = nothing
	Set ROService = nothing
	Set ROServer  = nothing
    Set fs = nothing
else
    'Inicia a exibição da imagem
	Response.Buffer = True
	Response.AddHeader "Content-Type","image/jpeg"
	Response.Flush
			
	Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.Open
    objStream.Type = adTypeText
    objStream.WriteText ""
    Response.BinaryWrite objStream.Read

	objStream.Close
	Set objStream = Nothing
	Response.Flush
end if
%>