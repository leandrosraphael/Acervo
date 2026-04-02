<%

On Error Resume Next
Set ROService = ROServer.CreateService("Web_Consulta")
sSubloc = ROService.DadosTabela(roTAB_SUBLOCALIZACAO,"",0,repositorio_institucional)
	
Set xmlSubloc = CreateObject("Microsoft.xmldom")
xmlSubloc.async = False
xmlSubloc.loadxml sSubloc
		
Set xmlRoot = xmlSubloc.documentElement
sListaSubloc = ""
'Verifica se a tabela foi encontrada e se possui registros
if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then

	For Each xmlTabela In xmlRoot.childNodes				
		if xmlTabela.nodeName  = "REGISTRO" then							
			sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
			sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
					
			sListaSubloc = sListaSubloc&"<option value='"&sCod&"'>"&sDesc&"</option>"	
		end if
	Next

else
	if (repositorio_institucional = 0) then
		Response.write "Erro na tabela Sublocalização!!!!"
		Response.end
	end if
end if
				
Set xmlSubloc    = nothing
Set ROService = nothing

estiloDiv = "display: none;"

if ((Request.QueryString("content") <> "autoridades") and (Request.QueryString("content") <> "autoridades_fak") and (GetModo_Busca <> "legislacao")) then
	estiloDiv = ""
end if
	
if (sListaSubloc <> "") or ((repositorio_institucional = 1) and (sListaSubloc = "")) then
%>
<div id="divSelectSubloc" style="display: inline;<%= estiloDiv %>">
	<select class="select_biblioteca" name="geral_subloc" id="geral_subloc" multiple="multiple" onKeyPress="return validaTecla(this,event,1,<%= global_numero_serie %>)" onChange="atualizaGeralSubloc();">
	<%
		Response.Write sListaSubloc
	%>
	</select>
	<input type="hidden" id="geral_subloc_codigos" name="geral_subloc_codigos"/>
</div>
<script>
	$(function () {
	<% if global_hab_bus_subloc = 1 and (global_espanha = 1 or (global_multibib = 1 or global_numero_serie = 5516 or global_numero_serie = 4516 or global_numero_serie = 7273)) and config_multi_servbib = 0 then %>
    
        
		$("#geral_subloc").multiselect({
			checkAllText: '',
			uncheckAllText: '',
			noneSelectedText: '<%= (getTermo(global_idioma, 7454, "Qualquer", 0) & " " & LCase(global_desc_subloc)) %>',
			selectedText: '# <%= getTermo(global_idioma, 4356, "Itens selecionados", 2) %>',
			selectedList: 1,
			header: false
		});		
			
	<% end if %>
	});
</script>
<%  end if %>