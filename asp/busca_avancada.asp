<%
Sub ComboFiltro(nome,def)
%>
	&nbsp;
	<select class='select_busca styled_combo' id="<%=nome%>" name="<%=nome%>" onkeypress="return validaTecla(1,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" onchange="atualizaComb_filtros();">
	<option value='palavra_chave' ><%=getTermo(global_idioma, 1354, "Todos os campos", 0)%></option>
	<option value='titulo'><%=getTermo(global_idioma, 177, "Título", 0)%></option>
	<option value='autor' >
		<% 
		   '**************************************************************************
		   ' Adequação IMS: Alteração de rótulos
		   '**************************************************************************
		   if global_numero_serie = 4516 then
			   Response.Write "Autor / Compositor"
		   else
			   Response.Write getTermo(global_idioma, 61, "Autor", 0)
		   end if
		%></option>
	<% 'Adequação IMS: Autores com funções específicas
	   if global_numero_serie = 4516 then %>
			<option value="autor_interprete">Intérprete</option>
			<option value="autor_acompanhante">Acompanhamento</option>
	<% end if %>
	<option value='assunto' ><%=getTermo(global_idioma, 72, "Assunto", 0)%></option>
	<%
	'**************************************************************************	
	' Adequação: Não exibe busca por ISBN e Série e Editora (BND)
	'**************************************************************************	
	if (global_numero_serie <> 5516) then %>
		<option value='editora' >
			<% 'Adequação IMS: Alteração de rótulos
			   if global_numero_serie = 4516 then
				   Response.Write "Editora / Gravadora"
			   else
				   Response.Write getTermo(global_idioma, 528, "Editora", 0)
			   end if
			%>
		</option>

		<option value='isbn_issn' ><%=getTermo(global_idioma, 1128, "ISBN", 0)%> / <%=getTermo(global_idioma, 1355, "ISSN", 0)%></option>
		<option value='serie' ><%=getTermo(global_idioma, 6897, "Série", 0)%></option>
	<% end if %>

	<% if global_versao = vSOPHIA then
		'**************************************************************************
		' Exibir o campo DGM de acordo com configuração
		'**************************************************************************
		if (global_busca_descricao_complementar = 1) then %>
			<option value='dgm' ><%=getTermo(global_idioma, 1460, "Desc. compl.", 0)%></option>
		<% end if

		if (global_pesquisar_fonte_avulsa = 1) then
		%>
			<option value="fonte"><%=getTermo(global_idioma, 180, "Fonte", 0)%></option>
		<%
		end if
	end if %>
	<!--<option value='idioma' >Idioma</option>-->
	<!--<option value='resumo' >Resumo</option>-->
	</select>
<%End Sub

Sub ComboLogica(nome)
%>
	<select class='select_logica styled_combo' id="<%=nome%>" name="<%=nome%>" onkeypress="return validaTecla(1,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" onchange="atualizaComb_logica();">
	<option value='E'><%=getTermo(global_idioma, 985, "E", 1)%></option>
	<option value='OU'><%=getTermo(global_idioma, 662, "OU", 1)%></option>
	<option value='NAO'><%=getTermo(global_idioma, 1356, "E NÃO", 1)%></option>
	</select>
<%
End Sub

'**********************
' TABELAS
'**********************
if config_multi_servbib <> 1 then
	On Error Resume Next
	
	Set ROService = ROServer.CreateService("Web_Consulta")	
	
	sXML_MATERIAL = ObterCache("tabMaterial_" & global_numero_serie, 5)

	if (sXML_MATERIAL = "") then
		sXML_MATERIAL = ROService.DadosTabela(roTAB_MATERIAL,"",0, repositorio_institucional)
		DefinirCache "tabMaterial_" & global_numero_serie, sXML_MATERIAL
	end if

	sXML_IDIOMAS = ObterCache("tabIdioma_" & global_numero_serie, 5)

	if (sXML_IDIOMAS = "") then
		sXML_IDIOMAS = ROService.DadosTabela(roTAB_IDIOMAS,"",0, repositorio_institucional)
		DefinirCache "tabIdioma_" & global_numero_serie, sXML_IDIOMAS
	end if

    sXML_REPOSDIGITAL = ObterCache("tabReposDigital_" & global_numero_serie, 5)

	if (sXML_REPOSDIGITAL = "") then
		sXML_REPOSDIGITAL = ROService.DadosTabela(roTAB_REPOSITORIO_DIGITAL,"",0, repositorio_institucional)
		DefinirCache "tabReposDigital_" & global_numero_serie, sXML_REPOSDIGITAL
	end if

	if (global_numero_serie = 5775) then
		sXML_FORMA_REGISTRO = ROService.DadosTabela(roTAB_FORMAREGISTRO,"",0, repositorio_institucional)
	end if
	
	if global_arquivo = 1 then
		sXML_NIVEIS = ROService.DadosTabela(roTAB_NIVEL_DESCRICAO,"",0, repositorio_institucional)
	end if

	if global_numero_serie = 5516 then
		sXML_MEIO_FISICO = ROService.DadosTabela(roTAB_MEIO_FISICO,"",0, repositorio_institucional)
		sXML_COLECAO = ROService.DadosTabela(roTAB_SUBLOCALIZACAO,"",0, repositorio_institucional)
	end if
	
	TrataErros(1)
	
	Set ROService = nothing
end if

%>
	<td>
	<form name="frm_combinada" action="index.asp?modo_busca=combinada&content=resultado&Servidor=<%=servidor%>&iBanner=<%=global_tipo_banner%>&iEscondeMenu=<%=global_esconde_menu%>&iIdioma=<%=global_idioma%>" method="post">	
	<tr>
	<% if global_numero_serie = 5516 then %>
		<td class="td_center_middle background_aba_ativa td_busca_avancada1" style="height: 105px;padding-top:10px;vertical-align:top;">
	<% else %>
		<td class="td_center_middle background_aba_ativa td_busca_avancada1" style="padding-top:10px;vertical-align: top;">
	<%
	end if
	input_class="input_busca"
	%>
	
	<table class="table-filtros" style="border-spacing: 2px; padding: 0;">
	<tr><td><% ComboFiltro "comb_filtro1","palavra_chave" %></td><td> <input class="<%=input_class%>" type="text" name="comb_campo1" onkeypress="return validaTecla(1,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" value='' onchange="atualizaComb_campos();" maxlength="500" /></td><td><% ComboLogica "comb_logica1" %></td></tr>
	<tr><td><% ComboFiltro "comb_filtro2","titulo" %></td><td> <input class="<%=input_class%>" type="text" name="comb_campo2" onkeypress="return validaTecla(1,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" value='' onchange="atualizaComb_campos();" maxlength="500" /></td><td><% ComboLogica "comb_logica2" %></td></tr>
	<tr><td><% ComboFiltro "comb_filtro3","autor" %></td><td> <input class="<%=input_class%>" type="text" name="comb_campo3" onkeypress="return validaTecla(1,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" value='' onchange="atualizaComb_campos();" maxlength="500" /></td><td><% ComboLogica "comb_logica3" %></td></tr>
	<tr><td><% ComboFiltro "comb_filtro4","assunto" %></td><td> <input class="<%=input_class%>" type="text" name="comb_campo4" onkeypress="return validaTecla(1,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" value='' onchange="atualizaComb_campos();" maxlength="500" /></td><td>&nbsp;</td></tr>
		
    <% if (repositorio_institucional = 1) and (global_bibdigital = 1) and (global_cfg_usa_repos_digital = 1) then %>
        <tr>
            <td class="td_ultimas_aquisicoes_busca ultimas_aquisicoes_busca_label"><%=getTermo(global_idioma, 270, "Repositório digital", 0)%></td>
            <td>
                <%
			    if (left(sXML_REPOSDIGITAL,5) = "<?xml") then
				    'Processa o XML
				    Set xmlDoc = CreateObject("Microsoft.xmldom")
				    xmlDoc.async = False
				    xmlDoc.loadxml sXML_REPOSDIGITAL
				    Set xmlRoot = xmlDoc.documentElement
		
				    'Verifica se a tabela foi encontrada e se possui registros
					Response.Write "<select name=""comb_reposdigital"" id=""comb_reposdigital"" multiple=""multiple"" class=""select_reposdigital"" onKeyPress=""return validaTecla(1,this,event,3,"&global_numero_serie&",'"&global_content&"','"&global_modo_busca&"',0);"" onChange=""atualizaCombReposDigital_opcoes();"">"
				    if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then    
                        
					    For Each xmlTabela In xmlRoot.childNodes				
						    if xmlTabela.nodeName  = "REGISTRO" then							
							    sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
							    sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
					
							    Response.write "<option value='"&sCod&"'>"&sDesc&"</option>"
						    end if
					    Next
					else
                        Response.Write "<option value=0 selected>"&getTermo(global_idioma, 1357, "Qualquer", 0)&"</option>"	
				    end if
                    Response.Write "</select><input type=""hidden"" id=""comb_reposdigital_codigos"" name=""comb_reposdigital_codigos""/>"
				
				    Set xmlRoot = nothing
				    Set xmlDoc = nothing
			    else
				    Response.write "Erro na tabela Repositório digital."
				    Response.end
			    end if
                %>
            </td>
        </tr>
    <%end if %>

	<% if (global_ultimas_aquisicoes = 1) then %>
		<tr>
			<% if (repositorio_institucional = 0) then %>
				<td class="td_ultimas_aquisicoes_busca ultimas_aquisicoes_busca_label"><%=getTermo(global_idioma, 7045, "Últimas aquisições", 0)%></td>
			<% else %>
				<td class="td_ultimas_aquisicoes_busca ultimas_aquisicoes_busca_label"><%=getTermo(global_idioma, 9236, "Últimos documentos", 0)%></td>
			<% end if %>
			<td class="td_ultimas_aquisicoes_busca_campos ultimas_aquisicoes_busca_data">
				<select tabindex="6" id="sel_data_aq" name="sel_data_aq" class="select_data styled_combo" onchange="habilitaEntre('aq');atualizaComb_campos();">
					<option value="0"><%=getTermo(global_idioma, 1073, "igual a", 2)%></option>
					<option value="1"><%=getTermo(global_idioma, 1074, "menor que", 2)%></option>
					<option value="2"><%=getTermo(global_idioma, 1075, "maior que", 2)%></option>
					<option value="3"><%=getTermo(global_idioma, 1076, "entre", 2)%></option>
				</select>

				<input type="text" maxlength="10" class="input_data" id="data_aq_inicio" name="data_aq_inicio" onkeypress="return formataCampoDatePicker(this, event)" onChange="atualizaComb_campos(); validarCampoData(this);" />
				<span id="data_aq_a" class="invisible"><label style="display:inline-block; width: 18px; text-align: center;"><%=getTermo(global_idioma, 1083, "a", 2)%></label></span>
				<span id="data_aq_comp" class="invisible">
					<input tabindex="8" type="text" maxlength="10" class="input_data" id="data_aq_fim" name="data_aq_fim" onkeypress="return formataCampoDatePicker(this, event)" onChange="atualizaComb_campos(); validarCampoData(this);" />
				</span>
			</td>
		</tr>
		<script type="text/javascript">
			$(function () {
				$.datepicker.setDefaults($.datepicker.regional['pt-BR']);
				$("#data_aq_inicio").datepicker({
					showOn: "both",
					buttonImage: "scripts/css/images/calendario.png",
					buttonImageOnly: true,
					minDate: new Date(1753, 1 - 1, 1),
					buttonText: ""
				});
				$("#data_aq_fim").datepicker({
					showOn: "both",
					buttonImage: "scripts/css/images/calendario.png",
					buttonImageOnly: true,
					minDate: new Date(1753, 1 - 1, 1),
					buttonText: ""
				});
		
				// Espanhol
				<% if (global_idioma = 1) then %>
					$("#data_aq_inicio").datepicker("option", $.datepicker.regional["es"]);
				    $("#data_aq_fim").datepicker("option", $.datepicker.regional["es"]);
				// Inglês
				<% elseif (global_idioma = 2) then %>
					$("#data_aq_inicio").datepicker("option", $.datepicker.regional[""]);
				    $("#data_aq_inicio").datepicker("option", "dateFormat", "dd/mm/yy");
			    	$("#data_aq_fim").datepicker("option", $.datepicker.regional[""]);
                    $("#data_aq_fim").datepicker("option", "dateFormat", "dd/mm/yy");
                // Catalão
                <% elseif (global_idioma = 3) then %>
                    $("#data_aq_inicio").datepicker("option", $.datepicker.regional["ca"]);
			    	$("#data_aq_fim").datepicker("option", $.datepicker.regional["ca"]);
                // Galego
                <% elseif (global_idioma = 5) then %>
                    $("#data_aq_inicio").datepicker("option", $.datepicker.regional["gl"]);
			    	$("#data_aq_fim").datepicker("option", $.datepicker.regional["gl"]);
                // Euskera
                <% elseif (global_idioma = 6) then %>
                    $("#data_aq_inicio").datepicker("option", $.datepicker.regional["eu"]);
			    	$("#data_aq_fim").datepicker("option", $.datepicker.regional["eu"]);
				<% end if %>

			});
		</script>
	<%end if %>

	<% if (global_numero_serie = 5775) then  %>
		<tr>
			<td style="width: 133px; text-align: left; padding-left:10px;"><%=getTermo(global_idioma, 7474, "Forma do registro", 0)%></td>
			<td style="text-align: left; padding-left: 5px;">
		<%
			
			if (left(sXML_FORMA_REGISTRO,5) = "<?xml") then
				'Processa o XML
				Set xmlDoc = CreateObject("Microsoft.xmldom")
				xmlDoc.async = False
				xmlDoc.loadxml sXML_FORMA_REGISTRO
				Set xmlRoot = xmlDoc.documentElement
		
				'Verifica se a tabela foi encontrada e se possui registros
				if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
					Response.Write "<select name=""comb_formaregistro"" id=""comb_formaregistro"" multiple=""multiple"" class=""select_avancada"" style=""width: 312px;"" onKeyPress=""return validaTecla(1,this,event,3,"&global_numero_serie&",'"&global_content&"','"&global_modo_busca&"',0);"" onChange=""atualizaCombFormaRegistro_opcoes();"">"
					
					For Each xmlTabela In xmlRoot.childNodes				
						if xmlTabela.nodeName  = "REGISTRO" then							
							sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
							sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
					
							Response.write "<option value='"&sCod&"'>"&sDesc&"</option>"
						end if
					Next
					
					Response.Write "</select><input type=""hidden"" id=""comb_formaregistro_codigos"" name=""comb_formaregistro_codigos""/>"
				else
					Response.write "Erro na tabela Forma do registro!!!!"
					Response.end
				end if
				
				Set xmlRoot = nothing
				Set xmlDoc = nothing
			else
				Response.write "Erro na tabela Forma do registro!!!!"
				Response.end
			end if
		%>
			</td>
		</tr>
	<%end if %>
	</table>
	</td>
	<td class="td_center_middle background_aba_ativa td_busca_avancada2" style="padding-top:10px;vertical-align: top;">
	<table class="table-filtros" style="border-spacing: 2px; padding: 0; display: inline-block;">
	<tr>
	<td class="td-label" style="width: 77px">
	  <% 'Adequação IMS: Alteração de rótulos
		 if global_numero_serie = 4516 then
			 Response.Write "Ano pub/grav"
		 else
			 Response.Write getTermo(global_idioma, 1358, "Ano edição", 0)
		 end if
	  %>
	</td>
	<td>
	<input class="input_ano" onkeyup="return validaTecla(1,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" type="text" name="comb_ano1" value='' onchange="atualizaComb_opcoes();" maxlength="500" />
	<% 'Adequação UNICAMP: Apenas um campo de ano, para fazer busca usando LIKE
	if global_numero_serie <> 4134 then %>
		<label class="label_entre_anos"><%=getTermo(global_idioma, 1083, "a", 2)%></label>
		<input class="input_ano" onkeyup="return validaTecla(1,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" type="text" name="comb_ano2" value='' onchange="atualizaComb_opcoes();" maxlength="500" />
	<% end if %>
	&nbsp;</td></tr>
	
	<% if (global_numero_serie = 5516 and global_hab_bus_subloc = 1) then  %>
		<tr><td class="td-label">Coleção</td><td>
		<%
			if (left(sXML_MATERIAL,5) = "<?xml") then
				'Processa o XML
				Set xmlDoc = CreateObject("Microsoft.xmldom")
				xmlDoc.async = False
				xmlDoc.loadxml sXML_COLECAO
				Set xmlRoot = xmlDoc.documentElement
		
				'Verifica se a tabela foi encontrada e se possui registros
				if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
					Response.Write "<select name=""comb_colecao"" id=""comb_colecao"" multiple=""multiple"" class=""select_avancada"" onKeyPress=""return validaTecla(1,this,event,3,"&global_numero_serie&",'"&global_content&"','"&global_modo_busca&"',0);"" onChange=""atualizaCombColecao_opcoes();"">"
					Response.Write "<option value=0 selected>"&getTermo(global_idioma, 1357, "Qualquer coleção", 0)&"</option>"
					For Each xmlTabela In xmlRoot.childNodes				
						if xmlTabela.nodeName  = "REGISTRO" then							
							sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
							sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
					
							Response.write "<option value='"&sCod&"'>"&sDesc&"</option>"
						end if
					Next
					
					Response.Write "</select><input type=""hidden"" id=""comb_colecao_codigos"" name=""comb_colecao_codigos""/>"
				else
					Response.write "Erro na tabela Coleção!!!!"
					Response.end
				end if
				
				Set xmlRoot = nothing
				Set xmlDoc = nothing
			else
				Response.write "Erro na tabela Coleção!!!!"
				Response.end
			end if
		%>
		</td></tr>

	<% end if %>

	<% if config_multi_servbib <> 1 then
		if global_numero_serie = 5516 then %>
			<tr><td class="td-label">Acervo</td><td>
			<%
				if (left(sXML_MEIO_FISICO,5) = "<?xml") then
					'Processa o XML
					Set xmlDoc = CreateObject("Microsoft.xmldom")
					xmlDoc.async = False
					xmlDoc.loadxml sXML_MEIO_FISICO
					Set xmlRoot = xmlDoc.documentElement
		
					'Verifica se a tabela foi encontrada e se possui registros
					if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
						Response.Write "<select name=""comb_meiofisico"" id=""comb_meiofisico"" multiple=""multiple"" class=""select_avancada"" onKeyPress=""return validaTecla(1,this,event,3,"&global_numero_serie&",'"&global_content&"','"&global_modo_busca&"',0);"" onChange=""atualizaCombMeioFisico_opcoes();"">"
					
						For Each xmlTabela In xmlRoot.childNodes				
							if xmlTabela.nodeName  = "REGISTRO" then							
								sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
								sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
					
								Response.write "<option value='"&sCod&"'>"&sDesc&"</option>"
							end if
						Next
					
						Response.Write "</select><input type=""hidden"" id=""comb_meiofisico_codigos"" name=""comb_meiofisico_codigos""/>"
					else
						Response.write "Erro na tabela Meio físico!!!!"
						Response.end
					end if
				
					Set xmlRoot = nothing
					Set xmlDoc = nothing
				else
					Response.write "Erro na tabela Meio físico!!!!"
					Response.end
				end if
			%>
			</td></tr>
		<% 
		end if
		%>
		<tr><td class="td-label"><%=getTermo(global_idioma, 175, "Material", 0)%></td><td>
		<%
			if (left(sXML_MATERIAL,5) = "<?xml") then
				'Processa o XML
				Set xmlDoc = CreateObject("Microsoft.xmldom")
				xmlDoc.async = False
				xmlDoc.loadxml sXML_MATERIAL
				Set xmlRoot = xmlDoc.documentElement
		
				'Verifica se a tabela foi encontrada e se possui registros
				Response.Write "<select name=""comb_material"" id=""comb_material"" multiple=""multiple"" class=""select_avancada"" onKeyPress=""return validaTecla(1,this,event,3,"&global_numero_serie&",'"&global_content&"','"&global_modo_busca&"',0);"" onChange=""atualizaCombMaterial_opcoes();"">"
				if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
					
					For Each xmlTabela In xmlRoot.childNodes				
						if xmlTabela.nodeName  = "REGISTRO" then							
							sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
							sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
					
							Response.write "<option value='"&sCod&"'>"&sDesc&"</option>"
						end if
					Next
				else
					Response.Write "<option value=0 selected>"&getTermo(global_idioma, 1357, "Qualquer", 0)&"</option>"
				end if
				Response.Write "</select><input type=""hidden"" id=""comb_material_codigos"" name=""comb_material_codigos""/>"
				
				Set xmlRoot = nothing
				Set xmlDoc = nothing
			else
				Response.write "Erro na tabela Midia!!!!"
				Response.end
			end if
		%>
		</td></tr>
		<% 
		'Adequação IMS: Busca por biblioteca / música
		if (repositorio_institucional = 0) and (global_numero_serie = 4516) then 
		%>
			<tr><td>&nbsp;</td>
			<td>
				<input type="checkbox" name="busca_musica" value="1" onchange="atualizaComb_opcoes();" /> Música
				&nbsp;
				<input type="checkbox" name="busca_biblioteca" value="1" onchange="atualizaComb_opcoes();" /> Biblioteca
				<input type="hidden" name="comb_idioma" value="0" />
			</td>
			</tr>
		<% else %>
			<tr><td class="td-label"><%=getTermo(global_idioma, 227, "Idioma", 0)%></td><td>
			<%
				if (left(sXML_IDIOMAS,5) = "<?xml") then
					'Processa o XML
					Set xmlDoc = CreateObject("Microsoft.xmldom")
					xmlDoc.async = False
					xmlDoc.loadxml sXML_IDIOMAS
					Set xmlRoot = xmlDoc.documentElement
			
					Response.Write "<select name=""comb_idioma"" id=""comb_idioma"" class=""select_avancada styled_combo"" onKeyPress=""return validaTecla(1,this,event,3,"&global_numero_serie&",'"&global_content&"','"&global_modo_busca&"',0);"" onChange=""atualizaComb_opcoes();"">"
					Response.Write "<option value=0 selected>"&getTermo(global_idioma, 1357, "Qualquer", 0)&"</option>"
					if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
						
						For Each xmlTabela In xmlRoot.childNodes				
							if xmlTabela.nodeName  = "REGISTRO" then							
								sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
								sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
						
								Response.write "<option value='"&sCod&"'>"&sDesc&"</option>"
							end if
						Next
					end if
					Response.Write "</select>"

					Set xmlRoot = nothing
					Set xmlDoc = nothing
				else
					Response.write "Erro na tabela Idiomas!!!!"
					Response.end
				end if
			%>
			</td></tr>
		<% end if %>

		<% if global_arquivo = 1 then 
			if (left(sXML_NIVEIS,5) = "<?xml") then
				'Processa o XML
				Set xmlDoc = CreateObject("Microsoft.xmldom")
				xmlDoc.async = False
				xmlDoc.loadxml sXML_NIVEIS
				Set xmlRoot = xmlDoc.documentElement
			
				'Verifica se a tabela foi encontrada e se possui registros
				if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
					Response.Write "<tr><td class=""td-label"">" & getTermo(global_idioma, 9877, "Arquivo", 0) & "</td><td>" 
					Response.Write "<select name=""comb_nivel"" id=""comb_nivel"" multiple=""multiple"" class=""select_avancada"" onKeyPress=""return validaTecla(1,this,event,3,"&global_numero_serie&",'"&global_content&"','"&global_modo_busca&"',0);"" onChange=""atualizaCombNivel_opcoes();"">"
					
					if (global_espanha = 1) then
						Response.Write "<option value=""0"">"&getTermo(global_idioma, 9635, "Coleção completa", 0)&"</option>"
					else
						Response.Write "<option value=""0"">"&getTermo(global_idioma, 7481, "Apenas pertencentes ao arquivo", 0)&"</option>"
					end if	
					For Each xmlTabela In xmlRoot.childNodes				
						if xmlTabela.nodeName  = "REGISTRO" then							
							sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
							sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
						
							Response.write "<option value='"&sCod&"'>"&sDesc&"</option>"
						end if
					Next
						
					Response.Write "</select><input type=""hidden"" id=""comb_nivel_codigos"" name=""comb_nivel_codigos""/>"
					Response.Write "</td></tr>"
				else
					Response.write "<input type=""hidden"" id=""comb_nivel_codigos"" name=""comb_nivel_codigos""/>"
				end if
					
				Set xmlRoot = nothing
				Set xmlDoc = nothing
			else
				Response.write "Erro na tabela Nivel descrição!!!!"
				Response.end
			end if
		end if %>
	<% end if %>	
	<tr>
		<td class="td-label"><%=getTermo(global_idioma, 1359, "Ordenação", 0)%></td><td>
		<select class="select_avancada styled_combo" id="comb_ordenacao" name="comb_ordenacao" onkeypress="return validaTecla(0,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" onChange="atualizaComb_opcoes();">
		<%if global_numero_serie = 5613 then%>
			<option value="data"><%=getTermo(global_idioma, 187, "Data", 0)%> - <%=getTermo(global_idioma, 1361, "crescente", 2)%></option>
			<option value="dataDESC"><%=getTermo(global_idioma, 187, "Data", 0)%> - <%=getTermo(global_idioma, 1362, "decrescente", 2)%></option>
			<option value="titulo"><%=getTermo(global_idioma, 177, "Título", 0)%> - <%=getTermo(global_idioma, 1361, "crescente", 2)%></option>
			<option value="tituloDESC"><%=getTermo(global_idioma, 177, "Título", 0)%> - <%=getTermo(global_idioma, 1362, "decrescente", 2)%></option>
		<%else%>	
			<option value="titulo"><%=getTermo(global_idioma, 177, "Título", 0)%> - <%=getTermo(global_idioma, 1361, "crescente", 2)%></option>
			<option value="tituloDESC"><%=getTermo(global_idioma, 177, "Título", 0)%> - <%=getTermo(global_idioma, 1362, "decrescente", 2)%></option>
		<%end if%>
			<option value="autor"><%=getTermo(global_idioma, 61, "Autor", 0)%> - <%=getTermo(global_idioma, 1361, "crescente", 2)%></option>
			<option value="autorDESC"><%=getTermo(global_idioma, 61, "Autor", 0)%> - <%=getTermo(global_idioma, 1362, "decrescente", 2)%></option>
			<option value="assunto"><%=getTermo(global_idioma, 72, "Assunto", 0)%> - <%=getTermo(global_idioma, 1361, "crescente", 2)%></option>
			<option value="assuntoDESC"><%=getTermo(global_idioma, 72, "Assunto", 0)%> - <%=getTermo(global_idioma, 1362, "decrescente", 2)%></option>
			<% if (global_espanha = 1) then  %>
			   <option value="ano"><%=getTermo(global_idioma, 67, "Ano", 0)%> - <%=getTermo(global_idioma, 1361, "crescente", 2)%></option>
			   <option value="anoDESC"><%=getTermo(global_idioma, 67, "Ano", 0)%> - <%=getTermo(global_idioma, 1362, "decrescente", 2)%></option>
			<% end if %>
		</select>
		</td>
	</tr>
	<% if (global_versao = vSOPHIA) then %>
		<tr><td>&nbsp;</td>
			<td class ="td-check"> 
				<span id="buscaMidiaSpanComb">
					<input type="checkbox" name="busca_midia" value="1" onchange="atualizaComb_opcoes();" /><%=getTermo(global_idioma, 573, "Registros com conteúdo digital", 0)%> 
				</span>
				&nbsp;
			</td>
		</tr>
	<% end if %> 
		
	<% if (ClienteItau) then %>
		<tr>
			<td class="td-label">
				<%
					On Error Resume Next
					Set ROService = ROServer.CreateService("Web_Consulta")
					'------------------------------------------------------------------------
					'----------------            Tabela Opcional            -----------------
					'------------------------------------------------------------------------
					nome_do_campo = ROService.DescricaoTabOpc
					TrataErros(1)
			
					Response.Write nome_do_campo
				%>
			</td>
			<td>
				<select class='select_avancada styled_combo' id="comb_tabopc" name="comb_tabopc" onKeyPress="return validaTecla(1,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" onChange="atualizaComb_tabopc();">
					<option value='0'><%=getTermo(global_idioma, 1357, "Qualquer", 0)%></option>        
					<%	
						sXML_TABOPC = ROService.DadosTabela(roTAB_TABOPC,"",0,repositorio_institucional)
						TrataErros(1)
			
						if (left(sXML_TABOPC,5) = "<?xml") then
							'Processa o XML
							Set xmlDoc = CreateObject("Microsoft.xmldom")
							xmlDoc.async = False
							xmlDoc.loadxml sXML_TABOPC
							Set xmlRoot = xmlDoc.documentElement
		
							'Verifica se a tabela foi encontrada e se possui registros
							if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then				
								For Each xmlTabela In xmlRoot.childNodes				
									if xmlTabela.nodeName  = "REGISTRO" then							
										sDesc = xmlTabela.attributes.getNamedItem("DESCRICAO").value
										sCod  = xmlTabela.attributes.getNamedItem("CODIGO").value
					
										Response.write "<option value='"&sCod&"'>"&sDesc&"</option>"
									end if
								Next
					
								Response.Write "</select>"
							end if
				
							Set xmlRoot = nothing
							Set xmlDoc = nothing
						end if
			
						Set ROService = nothing
					%>        
				</select>
			</td>
		</tr>
	<% end if %>

	<% if config_multi_servbib = 1 then %>
		<tr><td colspan="2" style="height: 45px">&nbsp;</td></tr>
	<% end if %>
	</table>
	</td>
	<% if global_numero_serie = 5516 then %>
		<td class="td_center_top background_aba_ativa td_busca_avancada3" style="height: 105px;padding-top:10px;">
	<% else %>
		<td class="td_center_top background_aba_ativa td_busca_avancada3" style="padding-top:10px;">
	<%
	end if
	%>
	<input style="margin-bottom: 5px;" title="<%=getTermo(global_idioma, 1363, "Submeter busca", 0)%>" class="button_busca" name="bt_comb" type="button" onclick="return Confere(<%=global_numero_serie%>,3,'<%=global_content%>','<%=global_modo_busca%>')" value="<%=label_buscar%>" />
	<input title="<%=getTermo(global_idioma, 1364, "Limpar campos", 0)%>" class="button_busca" name="bt_limpar" type="button" onClick="Resetar();atribuiComb();" value="<%=label_limpar%>" />
	</td></tr>	
	</table>
	<input type="hidden" name="comb_bib" value="0" />
	<input type="hidden" name="comb_subloc" value="0" />
	<input type="hidden" name="submeteu" value="combinada" />
  
   <%if GetModo_Busca = "ultimas_aquisicoes"  then%>
		<script type="text/javascript">
			Resetar();
			atribuiComb();
			global_frame.modo_busca = 'combinada';

			var dHoje = new Date(); 
			var dInicio = new Date(); 
			dInicio.setDate(dInicio.getDate() - global_dias_ultimas_aquisicoes);
			var diaInicio = dInicio.getDate();
			var mesInicio = dInicio.getMonth() + 1;
			var anoInicio = dInicio.getFullYear();
			var sInicio = diaInicio+'/'+mesInicio+'/'+anoInicio;

			var diaHoje = dHoje.getDate();
			var mesHoje = dHoje.getMonth() + 1;
			var anoHoje = dHoje.getFullYear();
			var sHoje = diaHoje+'/'+mesHoje+'/'+anoHoje;

			$("#data_aq_inicio").val(sInicio);
			$("#data_aq_fim").val(sHoje);
			$("#sel_data_aq").val(3);

			global_frame.comb_data_aq_inicio = sInicio;
			global_frame.comb_data_aq_fim = sHoje;
			global_frame.comb_sel_data_aq = 3;

			frm_combinada.submit();
		</script>
	<%end if%>
</form>	</td>