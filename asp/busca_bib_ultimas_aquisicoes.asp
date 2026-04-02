<%
if config_multi_servbib = 1 then
	'Monta com para busca de instituições
	if (Servidores.ServList.Count > 0) then
		sListaSrv = ""

		for iContaSrv = 0 to Servidores.ServList.Count-1
			sSrvName = Servidores.ServList.Item(iContaSrv).Nome
			sCodSrv = CStr(iContaSrv+1)
            
	
			if sSrvName <> "UNDEFINED" then
			   sListaSrv = sListaSrv & "<option value='" & sCodSrv & "'>" & sSrvName & "</option>"
            end if
		Next
%>
		<select class="select_biblioteca_ultimas_aquisicao styled_combo" name="geral_bib_ultimas_aquisicoes" id="geral_bib_ultimas_aquisicoes" onKeyPress="return validaTecla(this,event,1,<%= global_numero_serie %>)" onChange="atualizaselecaoUltimasAquisicoes(this);">

<%
		Response.Write sListaSrv   
%>
		</select><input type="hidden" id="geral_bib_codigos_ultimas_aquisicoes" name="geral_bib_codigos_ultimas_aquisicoes"/>
<%
	end if
end if
%>

<script type="text/javascript">
	var iServidor = 1;

	if (sessionStorage.iIndexSrvAquisicao) {
		iServidor = sessionStorage.iIndexSrvAquisicao;
	} 

	var comboCidades = document.getElementById("geral_bib_ultimas_aquisicoes");
	comboCidades.selectedIndex = iServidor - 1;

</script>

