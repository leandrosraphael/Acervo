<table class="max_width max_height">
<tr>
    <td class="td_padrao td_center_top" style="display: block">
		<br />
		<span style='color: #000000; font-weight: bold'><%=getTermo(global_idioma, 9687, "Certidão negativa de débitos", 0)%></span>
		<br /><br />
		<%
			if Session("codigo_usuario") = "" then
				mensagens = getTermoHtml(global_idioma, 1461, "Para ter acesso aos serviços da biblioteca você deve estar logado no sistema.", 0)
				Response.Write "<table class='tab_mensagem' style='display: inline-table'><tr><td class='esquerda'>"&mensagens&"</td></tr></table><br />"
			else 	
				codigo_usuario = Session("codigo_usuario")
	
				if config_multi_servbib = 1 then
					iIndexSrv = Session("Servidor_Logado")

					if iIndexSrv = "" then
						iIndexSrv = 1
					end if

					'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
					%><!-- #include file="../libasp/updChannelProperty.asp" --><%
				end if	

				On Error Resume Next

				Set ROService = ROServer.CreateService("Web_Consulta")
				Set CertidaoNegativa = ROServer.CreateComplexType("TRetornoCertidaoNegativa")
                Set CertidaoNegativa = ROService.EmitirCertidaoNegativa(codigo_usuario)
                                        
                If CertidaoNegativa.Result = true then                    
                    iCodigoCertidao = CertidaoNegativa.Codigo                    
                    %>
                    <script type="text/javascript" src="scripts/jquery.fileDownload.js"></script>					
                    <script>
                    $(document).ready(function () {	                    
						var urlDownload = "asp/certidao_negativa_download.asp?codigo=<%=iCodigoCertidao %>";
						$.fileDownload(urlDownload)                                
							.fail(function () { 
                                alert(getTermo(global_frame.iIdioma, 9696, "Falha ao fazer download.", 0)); 
                            });
                    });
                    </script>
                    <%
                              
                    sMensagem = getTermoHtml(global_idioma, 9697, "Certidão emitida.", 0)                 
                else
                    sMensagem = CertidaoNegativa.Mensagem                                      
                end if	

                Response.Write "<table class='tab_mensagem' style='display: inline-table'><tr><td class='esquerda'>"&sMensagem&"</td></tr></table><br />"                                        
                
			end if                                                      

            Set ROService = nothing
            Set CertidaoNegativa = nothing
		%>
	</td>
</tr>
</table>