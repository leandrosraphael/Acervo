<%
if config_multi_servbib = 1 then
	'Monta com para busca de instituições
	if (Servidores.ServList.Count > 0) then
		sListaSrv = ""
				
		for iContaSrv = 0 to Servidores.ServList.Count-1
			sSrvName = Servidores.ServList.Item(iContaSrv).Nome
			sCodSrv = CStr(iContaSrv+1)
			
			if Servidores.ServList.Item(iContaSrv).BibList.Count = 0 then
				if iContaSrv = 0 then
					sListaSrv = sListaSrv & "<optgroup class='grupo_selecao_servidor_first' label=''>"
				else			
					sListaSrv = sListaSrv & "<optgroup class='grupo_selecao_servidor' label=''>"
				end if			
				sListaSrv = sListaSrv & "<option class='selecao_servidor' value='" & CStr(iContaSrv+1) & ".0"& "'>" & sSrvName & "</option>"
				sListaSrv = sListaSrv & "</optgroup>"
			else
				if iContaSrv = 0 then
					sListaSrv = sListaSrv & "<optgroup class='grupo_selecao_servidor_first' label='" & sSrvName & "'>" 
				else			
					sListaSrv = sListaSrv & "<optgroup class='grupo_selecao_servidor' label='" & sSrvName & "'>" 
				end if				
		
				for iContaBib = 0 to Servidores.ServList.Item(iContaSrv).BibList.Count-1
					sSrvName = Servidores.ServList.Item(iContaSrv).BibList.Item(iContaBib).Nome
					sCodSrv = CStr(iContaSrv+1) & "." & CStr(Servidores.ServList.Item(iContaSrv).BibList.Item(iContaBib).Codigo)

					if sCodSrv = null or sCodSrv = "" or sCodSrv = "UNDEFINED" then
						sCodSrv = "0"
					end if

					if sSrvName = "UNDEFINED" then
						sSrvName = ""
					end if
	
					sListaSrv = sListaSrv & "<option value='" & sCodSrv & "'>" & sSrvName & "</option>"
				Next

				sListaSrv = sListaSrv & "</optgroup>"
			end if
		Next
%>
		<select class="select_biblioteca" name="geral_bib" id="geral_bib" multiple="multiple" onKeyPress="return validaTecla(this,event,1,<%= global_numero_serie %>)" onChange="atualizaGeral();">
<%
		Response.Write sListaSrv   
%>
		</select><input type="hidden" id="geral_bib_codigos" name="geral_bib_codigos"/>
<%
	end if
else
		On Error Resume Next
		Set ROService = ROServer.CreateService("Web_Consulta")
		sXML = ROService.BuscaBiblioteca(global_IP_Local = 1, repositorio_institucional)

		Set xmlBib = CreateObject("Microsoft.xmldom")
		xmlBib.async = False
		xmlBib.loadxml sXML
		
		Set xmlRoot = xmlBib.documentElement
	
		if xmlRoot.nodeName = "Bibliotecas" then	
	
			iNumBib = xmlRoot.attributes.getNamedItem("QTDE").Value
		
			sListaBib = ""
		
			For Each xmlBiblioteca In xmlRoot.childNodes
		
				if xmlBiblioteca.nodeName = "Biblioteca" then
					sNome 	 = xmlBiblioteca.attributes.getNamedItem("NOME").Value
					sCodigo	 = xmlBiblioteca.attributes.getNamedItem("CODIGO").value
				end if
		
				if sCodigo = null or sCodigo = "" or sCodigo = "UNDEFINED" then
					sCodigo = "0"
				end if

				if sNome = "UNDEFINED" then
					sNome = ""
				end if
	
				sListaBib = sListaBib&"<option value='"&sCodigo&"'>"&sNome&"</option>"
			Next
		end if	
	
		Set xmlBib    = nothing
		Set ROService = nothing
	
		if iNumBib > 0 then
	%>
			<select class="select_biblioteca" name="geral_bib" id="geral_bib" multiple="multiple" onKeyPress="return validaTecla(this,event,1,<%= global_numero_serie %>)" onChange="atualizaGeral();">
	<%
			Response.Write sListaBib
	%>   	</select><input type="hidden" id="geral_bib_codigos" name="geral_bib_codigos"/>
	<%  else 
			if (repositorio_institucional = 1) then %>
				<select class="select_biblioteca" name="geral_bib" id="geral_bib" multiple="multiple" onKeyPress="return validaTecla(this,event,1,<%= global_numero_serie %>)" onChange="atualizaGeral();">
					<option value='0'>Qualquer biblioteca</option>
				</select><input type="hidden" id="geral_bib_codigos" name="geral_bib_codigos"/>
			<% else %>
				<input type='hidden' name='geral_bib' id='geral_bib' value=''/>
				<input type='hidden' name='geral_bib_codigos' id='geral_bib_codigos' value=''/>
			<% end if %>
	<%  end if
end if
%>
<script type='text/javascript'>
	if (parent.hiddenFrame != null) {
		if (parent.hiddenFrame.iBusca_Projeto > 0)  {
			if (document.frm_combinada != null) {
				if (document.frm_combinada.comb_bib != null) {
					document.frm_combinada.comb_bib.style.visibility = "hidden";
				}
			}
			if (document.frm_geral != null) {
				if (document.frm_geral.geral_bib != null) {
					document.frm_geral.geral_bib.style.visibility = "hidden";
				}
			}
			if (document.frm_rapida != null) {
				if (document.frm_rapida.rapida_bib != null) {
					document.frm_rapida.rapida_bib.style.visibility = "hidden";
				}
			}
			if (document.frm_aut != null) {
				if (document.frm_aut.aut_bib != null) {
					document.frm_aut.aut_bib.style.visibility = "hidden";
				}
			}
		}
		if ((parent.hiddenFrame.iFixarBibUsu == 1) || (parent.hiddenFrame.iFixarBib == 1)) {
			if (document.frm_combinada != null) {
				if (document.frm_combinada.comb_bib != null) {
					document.frm_combinada.comb_bib.disabled = true;
				}
			}
			if (document.frm_geral != null) {
				if (document.frm_geral.geral_bib != null) {
				    document.frm_geral.geral_bib.disabled = true;
				}
			}
			if (document.frm_rapida != null) {
				if (document.frm_rapida.rapida_bib != null) {
					document.frm_rapida.rapida_bib.disabled = true;
				}
			}
			if (document.frm_aut != null) {
				if (document.frm_aut.aut_bib != null) {
					document.frm_aut.aut_bib.disabled = true;
				}
			}
		}
	}
</script>