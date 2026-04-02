<% if (Request.QueryString("ajax") = "1") then %>
	<!-- #include file="../libasp/funcoes.asp" -->
    <!-- #include file="../libasp/montar_campos.asp" -->
	<!-- #include file="../config.asp" -->
	<!-- #include file="../libasp/roclient.asp" -->
<% end if 

if (sMsgErro <> "") then
	Response.Write sMsgErro
else

%>   
<div id="divTituloPesquisa">
    <span style="float: left">Home</span>
    <!-- #include file="../asp/botaoLogin.asp" -->
	<!-- #include file="../asp/ler_parametros_busca.asp" -->
</div>
<br><br>
Forneça uma ou mais palavras para a pesquisa
<%
	if (Request.QueryString("msg_erro") <> "") then
		Response.Write "<br><br>"
		Response.Write "<img src='imagens/erro.gif'>&nbsp;<b>" & Request.QueryString("msg_erro") & "</b>"
	end if
	
	codigoUsuario = ClNg(Session("codigo"))
%>
<form action="asp/resumo.asp?content=detalhe_resultados" name="frm_pesquisa" method="POST" id="frm_pesquisa" onSubmit="return ValidaForm();">
<table style="width: 500px">
<tr>
<td colspan="2">&nbsp;</td>
</tr>
<tr>
<td class="direita">Palavra-chave</td>
<td class="esquerda"><input type="text" name="dados" id="dados" style="width: 285px" value="<%=sDados%>"></td>
</tr>
<%
	xmlCampos = ROService.GetCamposPesquisa(codigoUsuario)
	sMsgErro = TrataErros

	if (sMsgErro <> "") then
		Response.Write sMsgErro
	else
		Response.Write FormataCamposBusca(xmlCampos,sCampo1,sCampo2,sCampo3,sCampo4,sCampo5,sCampo6,sCampo7,sCampo8, "1")

		xmlObjetos = ROService.GetMaterial(Session("codigo"))
		sMsgErro = TrataErros
	
		if (sMsgErro <> "") then
			Response.Write sMsgErro
		else
			if (xmlObjetos <> "") then
				Response.Write ComboObjetos(xmlObjetos,iObjeto)
			end if
		    
			xmlContextos = ROService.GetContexto
			sMsgErro = TrataErros
			
			if (sMsgErro <> "") then
				Response.Write sMsgErro
			else
				Response.Write ComboContextos(xmlContextos,iContexto)		
%>
<tr id="campos_ordenacao">
	<td class='direita'>Ordenação</td>
	<td class='esquerda'>
		<select name='campo_ordenacao'>
		</select>
	</td>
</tr>
<tr id="ordem_pesquisa">
	<td class='direita'>Ordem</td>
	<td class='esquerda'>
		<select name='campo_ordem'>
<%
		if (OrdemPesquisa = "0") then
			Response.Write "<option value='0' selected>Crescente</option>"
			Response.Write "<option value='1'>Decrescente</option>"
		else
			Response.Write "<option value='0'>Crescente</option>"
			Response.Write "<option value='1' selected>Decrescente</option>"
		end if
%>
		</select>
	</td>
</tr>

<tr>
<td class="direita">&nbsp;</td>
<td class="esquerda">
<% if (Request.QueryString("imagem") = "1") then %>
	<input type="checkbox" name="imagem" id="imagem" value="1" checked>&nbsp;Somente itens com imagem</td>
<% else %>
	<input type="checkbox" name="imagem" id="imagem" value="1">&nbsp;Somente itens com imagem</td>
<% end if %>
</tr>
<tr>
<td colspan="2">&nbsp;</td>
</tr>    
<tr>
<td style="text-align: center" colspan="2">
	<input type="submit" name="submit" size="35" value="Buscar">&nbsp;&nbsp;
	<input type="button" name="reset" size="35" value="Limpar" onClick="resetForm()">
</td>
</tr>
</table>
</form>
<%
				xmlDica = ROService.GetDica
				sMsgErro = TrataErros
		
				if (sMsgErro <> "") then
					Response.Write sMsgErro
				else
					Response.Write QuadroDica(xmlDica)
					Response.Write "<br><br>"
	
%>
<script>
	var edBusca = document.getElementById('dados');
	if (edBusca != null) {
		edBusca.focus();
	}
</script>

<%
				end if
			end if
		end if
	end if
end if

Set ROService = nothing
Set ROServer = nothing
%>
<script type="text/javascript">
	$(function () {
		AtualizarComboMaterialOrdenacao(0);
	});
</script>