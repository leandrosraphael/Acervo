<%
if (Request.QueryString("aut") = 1) then %>
	<tr>
	<td class="background_aba_ativa td_busca_rapida">
		&nbsp;&nbsp;
		<form name='frm_aut' action='index.asp?modo_busca=rapida&content=autoridades&aut=1&iBanner=<%=global_tipo_banner%>&iEscondeMenu=<%=global_esconde_menu%>&iIdioma=<%=global_idioma%>' method='post'>
			<select class="select_busca styled_combo" name="aut_filtro" id="aut_filtro" onkeypress="return validaTecla(1,this,event,4,<%=global_numero_serie%>)" onchange="atualizaAuts();">
			<option value="qualquer"><%=getTermo(global_idioma, 1357, "Qualquer", 0)%></option>
			<option value="100"><%=getTermo(global_idioma, 732, "Pessoa", 0)%></option>
			<option value="110"><%=getTermo(global_idioma, 62, "Instituição", 0)%></option>
			<option value="111"><%=getTermo(global_idioma, 63, "Evento", 0)%></option>
			<option value="130"><%=getTermo(global_idioma, 70, "Título uniforme", 0)%></option>
			<option value="148"><%=getTermo(global_idioma, 71, "Termo cronológico", 0)%></option>
			<option value="150"><%=getTermo(global_idioma, 742, "Termo tópico", 0)%></option>
			<option value="151"><%=getTermo(global_idioma, 1148, "Local geográfico", 0)%></option>
			<option value="155"><%=getTermo(global_idioma, 3707, "Termo de gênero e forma", 0)%></option>
			<option value="180"><%=getTermo(global_idioma, 75, "Subdivisão geral", 0)%></option>
			<option value="181"><%=getTermo(global_idioma, 76, "Subdivisão geográfica", 0)%></option>
			<option value="182"><%=getTermo(global_idioma, 77, "Subdivisão cronológica", 0)%></option>
			<option value="185"><%=getTermo(global_idioma, 3390, "Subdivisão de forma", 0)%></option>
			<!--<option value="resumo">Resumo</option>-->
			</select>
			<input type='checkbox' name='aut_iniciado_com' value='1' onChange='atualizaAuts();'/><%=getTermo(global_idioma, 1239, "Iniciado com", 0)%>&nbsp;
			<input class="input_busca" type="text" name="aut_campo" id="aut_campo" onkeypress="return validaTecla(1,this,event,4,<%=global_numero_serie%>)" value="" onchange="atualizaAuts();" maxlength="500" />
			<input class="button_busca" type="button" title="<%=getTermo(global_idioma, 1363, "Submeter busca", 0)%>" onclick="return Confere_aut(<%=global_numero_serie%>,1,'<%=global_content%>','<%=global_modo_busca%>')" value="<%=label_buscar%>" />
			<input class="button_busca" type="button" title="<%=getTermo(global_idioma, 1364, "Limpar campos", 0)%>" onclick="Reset_auts();atribuiAuts();" value="<%=label_limpar%>" />
			<input type="hidden" name="aut_bib" value="0" />
			<input type="hidden" name="submeteu" value="aut" />
		</form>
	</td>
	</tr>
<% else %>
	<tr>
	<td class="background_aba_ativa td_busca_rapida">
		<form name='frm_rapida' action='index.asp?modo_busca=rapida&content=resultado&iBanner=<%=global_tipo_banner%>&iEscondeMenu=<%=global_esconde_menu%>&iIdioma=<%=global_idioma%>' method='post'>
			&nbsp;&nbsp;
			<select class="select_busca styled_combo" name="rapida_filtro" id="rapida_filtro" onkeypress="return validaTecla(1,this,event,1,<%=global_numero_serie%>);" onchange="atualizaRapida();">
			<option value="palavra_chave"><%=getTermo(global_idioma, 1354, "Todos os campos", 0)%></option>
			<option value="titulo"><%=getTermo(global_idioma, 177, "Título", 0)%></option>
			<option value="autor">
				<% 'Adequação IMS: Alteração de rótulos
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
			<option value="assunto"><%=getTermo(global_idioma, 72, "Assunto", 0)%></option>
			<%
			'**************************************************************************	
			' Adequação: Não exibe busca por ISBN e Série e Editora (BND)
			'**************************************************************************	
			if (global_numero_serie <> 5516) then %>
				<option value="editora">
					<% 'Adequação IMS: Alteração de rótulos
						if global_numero_serie = 4516 then
							Response.Write "Editora / Gravadora"
						else
							Response.Write getTermo(global_idioma, 528, "Editora", 0)
						end if
					%>
				</option>
			
				<option value="isbn_issn"><%=getTermo(global_idioma, 1128, "ISBN", 0)%> / <%=getTermo(global_idioma, 1355, "ISSN", 0)%></option>
				<option value="serie"><%=getTermo(global_idioma, 6897, "Série", 0)%></option>
			<% end if
			
            if global_versao = vSOPHIA then 
				'**************************************************************************	
				' Exibir o campo DGM de acordo com configuração
				'**************************************************************************	
				if (global_busca_descricao_complementar = 1) then %>
					<option value="dgm"><%=getTermo(global_idioma, 1460, "Desc. compl.", 0)%></option>
				<% end if

                if (global_pesquisar_fonte_avulsa = 1) then
                %>
			        <option value="fonte"><%=getTermo(global_idioma, 180, "Fonte", 0)%></option>
                <%
                end if
			end if %>
			<!--<option value="idioma">Idioma</option>-->
			<!--<option value="resumo">Resumo</option>-->
			</select>
			<input class="input_busca" type="text" name="rapida_campo" onkeypress="return validaTecla(0,this,event,1,<%=global_numero_serie%>)" value="" onchange="atualizaRapida();" maxlength="500" />
			<input class="button_busca" type="button" title="<%=getTermo(global_idioma, 1363, "Submeter busca", 0)%>" onclick="return Confere(<%=global_numero_serie%>,1,'<%=global_content%>','<%=global_modo_busca%>')" value="<%=label_buscar%>" />
			<input class="button_busca" type="button" title="<%=getTermo(global_idioma, 1364, "Limpar campos", 0)%>" onclick="Reset_rapida();atribuiRapida();" value="<%=label_limpar%>" />
			<input type="hidden" name="rapida_bib" value="0" />
			<input type="hidden" name="rapida_subloc" value="0" />
			<input type="hidden" name="submeteu" value="rapida" />
			<%
				if (global_versao = vSOPHIA) then
					'4516 = IMS; 5516 = BNBD
					if (global_numero_serie <> 4516) and (global_numero_serie <> 5516) then
						Response.Write "<span id='buscaMidiaSpan'>"
						Response.Write "<input type='checkbox' name='busca_midia' value='1' onChange='atualizaRapida();' />" 
						Response.Write getTermo(global_idioma, 573, "Registros com conteúdo digital", 0)
						Response.Write "</span>"
					end if
				end if 
			%>

			<% 
			'Adequação IMS: Busca por biblioteca / música
			if (global_numero_serie = 4516) then 
			%>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="checkbox" name="busca_musica" value="1" onchange="atualizaRapida();" /> Música
				&nbsp;
				<input type="checkbox" name="busca_biblioteca" value="1" onchange="atualizaRapida();" /> Biblioteca
			<%        
			end if 
			%>
		</form>
	</td>
	</tr>
<% end if %>