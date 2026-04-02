<script type="text/javascript">
	$(function () {
		$.datepicker.setDefaults($.datepicker.regional['pt-BR']);
		$("#data_ass_inicio").datepicker({
			showOn: "both",
			buttonImage: "scripts/css/images/calendario.png",
			buttonImageOnly: true,
			minDate: new Date(1753, 1 - 1, 1),
			buttonText: ""
		});
		$("#data_ass_fim").datepicker({
			showOn: "both",
			buttonImage: "scripts/css/images/calendario.png",
			buttonImageOnly: true,
			minDate: new Date(1753, 1 - 1, 1),
			buttonText: ""
		});
		$("#data_pub_inicio").datepicker({
			showOn: "both",
			buttonImage: "scripts/css/images/calendario.png",
			buttonImageOnly: true,
			minDate: new Date(1753, 1 - 1, 1),
			buttonText: ""
		});
		$("#data_pub_fim").datepicker({
			showOn: "both",
			buttonImage: "scripts/css/images/calendario.png",
			buttonImageOnly: true,
			minDate: new Date(1753, 1 - 1, 1),
			buttonText: ""
		});

		// Espanhol
		<% if (global_idioma = 1) then %>
            $("#data_ass_inicio").datepicker("option", $.datepicker.regional["es"]);
		    $("#data_ass_fim").datepicker("option", $.datepicker.regional["es"]);
		    $("#data_pub_inicio").datepicker("option", $.datepicker.regional["es"]);
		    $("#data_pub_fim").datepicker("option", $.datepicker.regional["es"]);
		// Inglês
		<% elseif (global_idioma = 2) then %>
            $("#data_ass_inicio").datepicker("option", $.datepicker.regional[""]);
		    $("#data_ass_inicio").datepicker("option", "dateFormat", "dd/mm/yy");
		    $("#data_ass_fim").datepicker("option", $.datepicker.regional[""]);
		    $("#data_ass_fim").datepicker("option", "dateFormat", "dd/mm/yy");
		    $("#data_pub_inicio").datepicker("option", $.datepicker.regional[""]);
		    $("#data_pub_inicio").datepicker("option", "dateFormat", "dd/mm/yy");
		    $("#data_pub_fim").datepicker("option", $.datepicker.regional[""]);
		    $("#data_pub_fim").datepicker("option", "dateFormat", "dd/mm/yy");
		// Catalão
        <% elseif (global_idioma = 3) then %>
            $("#data_ass_inicio").datepicker("option", $.datepicker.regional["ca"]);
            $("#data_ass_fim").datepicker("option", $.datepicker.regional["ca"]);
            $("#data_pub_inicio").datepicker("option", $.datepicker.regional["ca"]);
			$("#data_pub_fim").datepicker("option", $.datepicker.regional["ca"]);
        // Galego
        <% elseif (global_idioma = 5) then %>
            $("#data_ass_inicio").datepicker("option", $.datepicker.regional["gl"]);
            $("#data_ass_fim").datepicker("option", $.datepicker.regional["gl"]);
            $("#data_pub_inicio").datepicker("option", $.datepicker.regional["gl"]);
			$("#data_pub_fim").datepicker("option", $.datepicker.regional["gl"]);
        // Euskera
        <% elseif (global_idioma = 6) then %>
            $("#data_ass_inicio").datepicker("option", $.datepicker.regional["eu"]);
            $("#data_ass_fim").datepicker("option", $.datepicker.regional["eu"]);
            $("#data_pub_inicio").datepicker("option", $.datepicker.regional["eu"]);
			$("#data_pub_fim").datepicker("option", $.datepicker.regional["eu"]);
		<% end if %>
	});
</script>

<td>
<form name="frm_legislacao" action="index.asp?modo_busca=legislacao&content=resultado&iBanner=<%=global_tipo_banner%>&iEscondeMenu=<%=global_esconde_menu%>&iSomenteLegislacao=<%=global_somente_legislacao%>&iIdioma=<%=global_idioma%>" method="post">
	<tr>
	<td class="background_aba_ativa td_busca_leg1 td_center_middle" style="padding-top: 6px;">
		
        <!-- Busca simplificada -->
        <table style="height: 84px;" class="remover_bordas_padding max_width">
	    <tr>
			<td colspan="2" class="label_norma_tj"><b><%=getTermo(global_idioma, 1173, "Norma", 0)%></b></td>
			<td colspan="2" class="label_orgao_origem_tj"><b><%=getTermo(global_idioma, 1369, "Órgão de origem", 0)%></b></td>
		</tr>	
		<tr>
			<td colspan="2">
				<div class="div_selecao_tj normas_tj" id="div_selecao_norma_tj">
				</div>
			</td>
			<td colspan="2" class="padding_orgao_origem_tj">
				<div class="div_selecao_tj orgao_origem_tj" id="div_origem_tj">
				</div>
			</td>
	    </tr>
		<%
			Response.Write "<input type='hidden' id='ckNorma_codigos' name='ckNorma_codigos' class='ckNorma_codigos'/>"
			Response.Write "<input type='hidden' id='ckOrgaoOrigem_codigos' name='ckOrgaoOrigem_codigos' class='ckOrgaoOrigem_codigos'/>"
		%>
	    </table>

		<table style="border-spacing: 2px" class="max_width">
			<tr>
				<td style="width: 140px; height: 30px">&nbsp;&nbsp;<%=getTermo(global_idioma, 1354, "Todos os campos", 0)%></td>
				<td><input tabindex="1" style="margin-top: 0px" class="inputLegEsq" type="text" name="leg_campo1" onkeypress="return validaTecla(1,this,event,5,<%=global_numero_serie%>,'<%=global_content%>','parent.hiddenFrame.modo_busca',0)" value='' onchange="atualizaLeg_campos();" maxlength="500" /></td>		
				<td style="width: 75px; padding-left: 10px;"><%=getTermo(global_idioma, 101, "Número", 0)%></td>
				<td>
					<input tabindex="3"  title="<%="Favor desconsiderar o zero a esquerda do Ato (Res. nº 06). Exemplo: 6"%>" style="width: 180px; margin-right: 3px;" type="text" name="leg_numero" onkeypress="return validaTecla(1,this,event,5,<%=global_numero_serie%>,'<%=global_content%>','parent.hiddenFrame.modo_busca',0)" value='' onchange="atualizaLeg_campos();" maxlength="500" />
					<label style="display:inline-block; width: 28px;"><%=getTermo(global_idioma, 67, "Ano", 0)%></label>
					<input tabindex="4" style="width: 91px;" type="text" name="ano_ass" onkeypress="return (BloqueiaNaoNumerico(event) && validaTecla(1,this,event,5,<%=global_numero_serie%>,'<%=global_content%>','parent.hiddenFrame.modo_busca',0));" value=''  onchange="atualizaLeg_campos();" maxlength="500" />
				</td>
			</tr>
	    </table>

        <!-- Busca Avançada -->
        <div id="div_leg_avancada" style="display: none">
            <table style="height: 140px; border-spacing: 2px" class="max_width">
	        <tr>
				<td style="width: 140px; height: 28px">&nbsp;&nbsp;<%=getTermo(global_idioma, 1319, "Ementa", 0)%></td>
                <td><input tabindex="9" class="inputLegEsq" type="text" name="leg_campo5" onkeypress="return validaTecla(1,this,event,5,<%=global_numero_serie%>,'<%=global_content%>','parent.hiddenFrame.modo_busca',0)" value='' onchange="atualizaLeg_campos();" maxlength="500" /></td>	
				
				<td style="width: 75px; padding-left: 16px;"><%=getTermo(global_idioma, 807, "Assinatura", 0)%></td>
                <td>
					<select tabindex="6" id="sel_data_ass" name="sel_data_ass" class="select_data styled_combo" onchange="habilitaEntre('ass');atualizaLeg_campos();">
			            <option value="0"><%=getTermo(global_idioma, 1073, "igual a", 2)%></option>
			            <option value="1"><%=getTermo(global_idioma, 1074, "menor que", 2)%></option>
			            <option value="2"><%=getTermo(global_idioma, 1075, "maior que", 2)%></option>
			            <option value="3"><%=getTermo(global_idioma, 1076, "entre", 2)%></option>
			        </select>

					<input tabindex="7" type="text" maxlength="10" class="input_data" id="data_ass_inicio" name="data_ass_inicio" onkeypress="return formataCampoDatePicker(this, event)" onChange="atualizaLeg_campos(); validarCampoData(this);" />
			        <span id="data_ass_a" class="invisible"><label style="display:inline-block; width: 30px; text-align: center;"><%=getTermo(global_idioma, 1083, "a", 2)%></label></span>
			        <span id="data_ass_comp" class="invisible">
						<input tabindex="8" type="text" maxlength="10" class="input_data" id="data_ass_fim" name="data_ass_fim" onkeypress="return formataCampoDatePicker(this, event)" onChange="atualizaLeg_campos(); validarCampoData(this);" />
					</span>
				</td>
	        </tr>
            <tr>
	            <td>&nbsp;&nbsp;<%=getTermo(global_idioma, 1320, "Texto integral", 0)%></td>
                <td><input tabindex="13" class="inputLegEsq" type="text" name="leg_campo6" onkeypress="return validaTecla(1,this,event,5,<%=global_numero_serie%>,'<%=global_content%>','parent.hiddenFrame.modo_busca',0)" value='' onchange="atualizaLeg_campos();" maxlength="500" /></td>
	            <td style="padding-left: 16px;"><%=getTermo(global_idioma, 3649, "Publicação", 0)%></td>
                <td>
					<select tabindex="10" id="sel_data_pub" name="sel_data_pub" class="select_data styled_combo" onchange="habilitaEntre('pub');atualizaLeg_campos();">
	                    <option value="0"><%=getTermo(global_idioma, 1073, "igual a", 2)%></option>
	                    <option value="1"><%=getTermo(global_idioma, 1074, "menor que", 2)%></option>
	                    <option value="2"><%=getTermo(global_idioma, 1075, "maior que", 2)%></option>
	                    <option value="3"><%=getTermo(global_idioma, 1076, "entre", 2)%></option>
	                </select>
					
					<input tabindex="11" type="text" maxlength="10" class="input_data" id="data_pub_inicio" name="data_pub_inicio" onkeypress="return formataCampoDatePicker(this, event)" onChange="atualizaLeg_campos(); validarCampoData(this);"/>
	                <span id="data_pub_a" class="invisible"><label style="display:inline-block; width: 30px; text-align: center;"><%=getTermo(global_idioma, 1083, "a", 2)%></label></span>
	                <span id="data_pub_comp" class="invisible">
						<input tabindex="12" type="text" maxlength="10" class="input_data" id="data_pub_fim" name="data_pub_fim" onkeypress="return formataCampoDatePicker(this, event)" onChange="atualizaLeg_campos(); validarCampoData(this);" />
					</span>
				</td>
	        </tr>
	        <tr>
	            <td>&nbsp;&nbsp;<%=getTermo(global_idioma, 205, "Assuntos", 0)%></td>
			    <td style="width: 307px">
					<input tabindex="15" class="inputLegEsq" type="text" name="leg_campo4" onkeypress="return validaTecla(1,this,event,5,<%=global_numero_serie%>,'<%=global_content%>','parent.hiddenFrame.modo_busca',0)" value='' onchange="atualizaLeg_campos();" maxlength="500" />
			    </td>
				<input type="hidden" class="inputLeg" type="text" name="processo" />

				<td colspan="2" rowspan="2">
					<fieldset class="fieldset-group">
						<legend><b><%=getTermo(global_idioma, 5954, "Projeto de lei", 0)%></b></legend>
						<table style="height: 25px" class="remover_bordas_padding max_width">
						<tr>
							<td style="width: 82px">&nbsp;&nbsp;<%=getTermo(global_idioma, 5955, "Autoria", 0)%></td>
							<td style="width: 201px"><input style="width: 180px" tabindex="17" class="inputLeg" type="text" name="leg_autoria" onKeyPress="return validaTecla(1,this,event,5,<%=global_numero_serie%>,'<%=global_content%>','parent.hiddenFrame.modo_busca',1)" value="" onChange="atualizaLeg_campos();" maxlength="500"></td>
							<td style="width: 22px"><%=getTermo(global_idioma, 186, "Nº", 0)%></td>
							<td><input style="width: 91px" tabindex="18" class="inputLeg" type="text" name="leg_numero_projeto" onKeyPress="return validaTecla(1,this,event,5,<%=global_numero_serie%>,'<%=global_content%>','parent.hiddenFrame.modo_busca',1)" value="" onChange="atualizaLeg_campos();" maxlength="500"></td>
						</tr>
						</table>
					</fieldset>
				</td>

	        </tr>
            <tr>
			    <td>&nbsp;&nbsp;<%=getTermo(global_idioma, 1370, "Resp. intelectual", 0)%></td>
	            <td>
					<input tabindex="16" class="inputLegEsq" type="text" name="leg_campo2" onkeypress="return validaTecla(1,this,event,5,<%=global_numero_serie%>,'<%=global_content%>','parent.hiddenFrame.modo_busca',0)" value='' onchange="atualizaLeg_campos();" maxlength="500" />
	            </td>
	        </tr>
            <tr>
				<td style="width: 140px; height: 28px"></td>
                <td class ="td-check">
      	           <span id="buscaMidiaSpanLeg">
					   <input type="checkbox" name="busca_midia" value="1" onchange="atualizaLeg_campos();" /><%=getTermo(global_idioma, 573, "Registros com conteúdo digital", 0)%> 
                    </span>&nbsp;
                </td>

				<td style="width: 60px;">&nbsp;&nbsp;&nbsp;&nbsp;<%=getTermo(global_idioma, 1359, "Ordenação", 0)%></td>
				<td>
				<select style="width: 316px;" class="styled_combo" id="leg_ordenacao" name="leg_ordenacao" onkeypress="return validaTecla(0,this,event,3,<%=global_numero_serie%>,'<%=global_content%>','<%=global_modo_busca%>',0)" onChange="atualizaLeg_campos();">
						<option value="titulo"><%=getTermo(global_idioma, 1173, "Norma", 0)%></option>
						<option value="publicacaoDESC"><%=getTermo(global_idioma, 10, "Data de publicação", 0)%> - <%=getTermo(global_idioma, 1362, "decrescente", 2)%></option>
						<option value="publicacao"><%=getTermo(global_idioma, 10, "Data de publicação", 0)%> - <%=getTermo(global_idioma, 1361, "crescente", 2)%></option>
					</select>
				</td>
	        </tr>
	        </table>
        </div>

        <!-- Separação da busca avançada -->
        <div style="margin-top: 16px; padding: 4px; padding-left: 95px">
            <a id="buscaAvancadaLegLink" tabindex="19" href="#" class="link_abas" style="font-weight: bold"><%=getTermo(global_idioma, 6895, "Exibir busca avançada", 0)%></a>
        </div>

	</td>
	<td class="background_aba_ativa td_center_top" style="vertical-align: middle;">
	    <input style="margin-bottom: 5px;" tabindex="20" title="<%=getTermo(global_idioma, 1363, "Submeter busca", 0)%>" class="button_busca" name="bt_comb" type="button" onclick="return Confere(<%=global_numero_serie%>,5,'<%=global_content%>',parent.hiddenFrame.modo_busca)" value="<%=label_buscar%>" />
	    <input tabindex="21" title="<%=getTermo(global_idioma, 1364, "Limpar campos", 0)%>" class="button_busca" name="bt_limpar" type="button" onClick="Resetar();atribuiLegs();" value="<%=label_limpar%>" />
	</td>
    </tr>	
	</table>
	<input type="hidden" name="leg_bib" value="0" />
	<input type="hidden" name="submeteu" value="legislacao" />
</form></td>

<script type="text/javascript">

</script>
