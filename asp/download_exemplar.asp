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
	Function getFileExt(sArquivo)
	  Dim arPath
	  arPath = Split(sArquivo, ".")
	  GetFileExt = arPath(UBound(arPath,1))
	End Function

	iIndexSrv = IntQueryString("iIndexSrv", 1)
	
%><!-- #include file="../libasp/updChannelProperty.asp" --><%	

	codEx = Request.QueryString("codEx")
	codMidia = Request.QueryString("codigo")
	tipoMidia = Request.QueryString("tipo_midia")
	video = Request.QueryString("video")
	sArquivo = ""
    usuario = Request.QueryString("iUsuario")
	
	On Error Resume Next

	'******************************************
	'Contagem de acesso
	'******************************************
	Set ROService = ROServer.CreateService("Web_Consulta")
	ROService.ContaAcessosMidiaExemplar CLng(codMidia), CLng(tipoMidia), CLng(usuario)
	Set ROService = nothing
	
	'******************************************
	'Só redireciona, midia concentrada
	'******************************************
	if (Request.QueryString("redirect") = "1") then
		Set ROService = ROServer.CreateService("Web_Consulta")
		ROService.MidiaDownload CLng(codMidia), CLng(Request.QueryString("iUsuario"))
		Set ROService = nothing
		
		if (err.number <> 0) then
			Response.Write "<div class='centro'><span style='color: #000033; font: Helvetica;'>"
			Response.Write "<span class='span_imagem icon_16 erro'></span>&nbsp;"
			if (Session("debug") = "habilitado") then
				Response.Write "<b>" & getTermo(global_idioma, 395, "Erro", 0) & " " &err.number&"</b> - "&err.description
			else
				Response.Write getTermo(global_idioma, 1457, "Mídia indisponível no momento.", 0)
			end if
			Response.Write "</span></div>"
		else
			sNomeArq = Request.QueryString("arq")
			
			Response.Redirect (global_midia_caminho_virtual & "/" & sNomeArq)
		end if
	'******************************************
	'Midias espalhadas, efetua o download
	'******************************************
	else	
		Set ROService = ROServer.CreateService("Web_Consulta")	
		
        'Solicita ao servidor o inicio do download
        'O arquivo será fragmentado no servidor e o mesmo retornará em quantas partes o arquivo será dividido
        Set DownloadMidia = ROService.InicializaDownloadMidiaExemplar(CLng(codMidia), CLng(tipoMidia))
        sNomeArq = DownloadMidia.NomeArquivo
		
		if (DownloadMidia.Partes <= 0) then
			'Servidor não encontrou a mídia
			Response.Write "<div class='centro'><span style='color: #000033; font-family:arial, helvetica;'>"
			Response.Write "<span class='span_imagem icon_16 erro'></span>&nbsp;"
			if (Session("debug") = "habilitado") then
				Response.Write getTermo(global_idioma, 6672, "O servidor de aplicação não encontrou a mídia.", 0)
			else
				Response.Write getTermo(global_idioma, 1457, "Mídia indisponível no momento.", 0)
			end if
			Response.Write "</span></div>"
				
			Set DownloadMidia = nothing
			Set ROService = nothing
			Set ROServer  = nothing
		else
			bStreaming = False
			if tipoMidia = "2" then
				bStreaming = ROService.HabilitaStreaming(CLng(codMidia))
			end if

            'Quantidade de partes
            TotalPartes = DownloadMidia.Partes
            'Identificador do download do arquivo
            Guid = DownloadMidia.GUID

            Set fs = Server.CreateObject("Scripting.FileSystemObject")

            ' Verifica se a pasta temp exista, caso não exista ela é criada
            If fs.FolderExists(Server.MapPath("../temp")) <> true Then
                fs.CreateFolder(Server.MapPath("../temp"))
            End if

            ' Cria a pasta com o identificador para armazenar o arquivo fragmentado
            PastaGuid = Server.MapPath("../temp") & "\" & Guid
            fs.CreateFolder(PastaGuid)

            ' Solicita todas as partes do arquivo
            For i = 1 to TotalPartes
                ' Solicita o arquivo para o servidor
                Set BinParte = ROService.DownloadParteMidia(Guid, i)

                if (BinParte.ToString <> "") then
                    BinParte.SaveToFile(PastaGuid & "\O" & codMidia & "._" & i)
                else
			        'Servidor não encontrou a parte
			        Response.Write "<div class='centro'><span style='color: 000033; font: font-family: arial, Helvetica'>"
			        Response.Write "<span class='span_imagem icon_16 erro'></span>&nbsp;"
			        if (Session("debug") = "habilitado") then
				        Response.Write getTermo(global_idioma, 6673, "Erro ao carregar parte da mídia.", 0)
			        else
				        Response.Write getTermo(global_idioma, 1457, "Mídia indisponível no momento.", 0)
			        end if
			        Response.Write "</span></div>"
				
			        Set BinParte = nothing
			        Set ROService = nothing
			        Set ROServer  = nothing
                    Set fs = nothing

                    Response.End
                end if

                Set BinParte = nothing
            Next

            ' Avisa o servidor que o download foi finalizado
            ' O servidor irá excluir os arquivos criados
            ROService.FinalizaDownloadMidia(Guid)

			sID = "O" & codMidia & "." & getFileExt(sNomeArq)
					
			' Se for streaming o arquivo é agrupado novamente
            if bStreaming OR CStr(video) = "1" then
                const adTypeBinary = 1
                const adSaveCreateOverwrite = 2

                Set objStreamOutput = Server.CreateObject("ADODB.Stream")
                objStreamOutput.Open
                objStreamOutput.Type = adTypeBinary

                For i = 1 to TotalPartes
					Set objStream = Server.CreateObject("ADODB.Stream")
						    
                    objStream.Open
					objStream.Type = adTypeBinary
					objStream.LoadFromFile PastaGuid & "\O" & codMidia & "._" & i
						    
                    objStreamOutput.Write objStream.Read
						    
                    objStream.Close
					Set objStream = Nothing
                Next

				if bStreaming then
					local_arquivo = global_streaming_caminho_fisico & "\" & sID
				else
					local_arquivo = Server.MapPath("../temp") & "\" & sID
				end if

				objStreamOutput.SaveToFile local_arquivo, adSaveCreateOverwrite

                objStreamOutput.Close
                objStreamOutput = nothing
				
				'Se ocorrer erro, verifica se o arquivo já existe
				'Pois se tentar sobrescrever o arquivo e o mesmo estiver em uso ocorre um erro
				if (err.number <> 0) and fs.FileExists(local_arquivo) then
					err.clear
				end if
			end if
		
			if (err.number <> 0) then
				'Ocorreu um erro ao gravar a mídia
				Response.Write "<div class='centro'><span style='color: #000033; font-family:arial, helvetica;'>"
				Response.Write "<span class='span_imagem icon_16 erro'></span>&nbsp;"
				if (Session("debug") = "habilitado") then
					Response.Write "<b>" & getTermo(global_idioma, 395, "Erro", 0) & " " & err.number&"</b> - "&err.description
				else
					Response.Write getTermo(global_idioma, 1457, "Mídia indisponível no momento.", 0)
				end if
				Response.Write "</span></div>"
			else
				ROService.MidiaDownload CLng(codMidia), CLng(Request.QueryString("iUsuario"))
			
				' Se for streaming redireciona para o Player
                if bStreaming OR CStr(video) = "1" then							
					'Envia POST			
					Response.Write "<html>"	
					Response.Write "<body onload=""document.getElementById('frmMidia').submit();"" >"
					Response.Write "<form id='frmMidia' method='POST' action='../player/SophiAPlayer.asp'>"
					Response.Write "<input type='hidden' name='midia' value='" & getTermo(global_idioma, 218, "Mídias", 0) & "'/>"
					if bStreaming then
						Response.Write "<input type='hidden' name='streaming' value='" & global_streaming_caminho_virtual & "/" & sID & "'/>"
					else
						Response.Write "<input type='hidden' name='video' value='../temp/" & sID & "'/>"
						Response.Write "<input type='hidden' name='extensao' value='" & Request.QueryString("extensao") & "'/>"
					end if
					Response.Write "<input type='hidden' name='obra' value='" & Request.QueryString("obra") & "'/>"
					Response.Write "<input type='hidden' name='tipo' value='" & Request.QueryString("tipo") & "'/>"
					Response.Write "<input type='hidden' name='iIndexSrv' value='" & iIndexSrv & "'/>"					
					Response.Write "</form>"
					Response.Write "</body>"
					Response.Write "<html>"
				else
					if (Session("debug") = "habilitado") then
						'Mídia encontrada porém o modo debug está habilitado
						Response.Write "<div class='centro'><span style='color: #000033; font-family:arial, helvetica;'>"
						Response.Write getTermo(global_idioma, 6674, "Mídia encontrada, mas não pode ser visualizada no modo debug.", 0)
						Response.Write "</span></div>"
					else
						'Inicia a exibição da mídia
						Response.Buffer = True
						Response.AddHeader "Content-Type","application/octet-stream"
						Response.AddHeader "Content-Disposition","attachment; filename=""" & sNomeArq & """"
						Response.Flush
						
                        ' Envia para o browser cada uma das partes do arquivo
                        For i = 1 to TotalPartes
						    Set objStream = Server.CreateObject("ADODB.Stream")
						    
                            objStream.Open
						    objStream.Type = 1
						    objStream.LoadFromFile PastaGuid & "\O" & codMidia & "._" & i
						    
                            Response.BinaryWrite objStream.Read
						    
                            objStream.Close
						    Set objStream = Nothing

						    Response.Flush
                        Next
					end if
				end if
			end if

            'Exclui a pasta com a mídia fragmentada
            if fs.FolderExists(PastaGuid) then
                fs.DeleteFile PastaGuid & "\*.*"
                fs.DeleteFolder PastaGuid, true
            end if

			Set DownloadMidia = nothing
            Set ROService = nothing
			Set ROServer = nothing
            Set fs = nothing
		end if
	end if
%>