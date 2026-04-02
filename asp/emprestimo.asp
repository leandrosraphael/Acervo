<table class="max_width max_height">
<tr>
<td class="td_padrao td_center_top" style="display: block">
<script type="text/javascript" src="scripts/emprestimos.js"></script>
<%
if Session("codigo_usuario") = "" then
	Response.Redirect("asp/logout.asp?logout=sim&modo_busca=" & GetModo_Busca & "&iBanner="&global_tipo_banner&"&iIdioma=" & global_idioma)
end if

if Session("nome_usuario") <> "" then
	%>

	<div id="frmEmprestimo">
		<%if global_cfg_exemplar_tombo = 1 then %>
			<label for="exemplar"><%=getTermo(global_idioma, 192, "Tombo", 0)%></label>
			<input type="text" name="exemplar" id="idexemplar" maxlength="20" onkeypress="javascript:return keyPressEmprestimo(event, 'tombo');" />
		<% else %>
			<label for="exemplar"><%=getTermo(global_idioma, 81, "Código", 0)%></label>
			<input type="text" name="exemplar" id="idexemplar" maxlength="10" onkeypress="javascript:return keyPressEmprestimo(event, 'codigo');" />
		<%end if %>
		<input type="hidden" name="idioma" value="<%= global_idioma %>" />
		<input type="hidden" name="usuario" value="<%=Session("codigo_usuario")%>" />
		<input type="hidden" name="tipo" value="-1" />

		<input type="submit" value="<%=getTermo(global_idioma, 444, "Buscar", 0)%>" onclick="javascript:emprestar()" />

		<div class="ficha-emprestimo"></div>

	</div>
<%else%>
	erro
<%end if%>
</td>
</tr>
</table>