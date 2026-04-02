<% 
	idioma = IntQueryString("idioma", 0)

	iIndexSrv = IntQueryString("servidor", 1)
	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%


	Set ROService = ROServer.CreateService("Web_Consulta")	
	sXML = ROService.ObterListaFavorito(Session("codigo_usuario"), idioma, repositorio_institucional)
	Set ROService = nothing
	TrataErros(1)

	if (left(sXML,5) = "<?xml") then
		'Processa o XML
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml sXML
		Set xmlRoot = xmlDoc.documentElement

		'Verifica se a tabela foi encontrada e se possui registros
		if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
			%>
			<table class="max_width max_height">
				<tr>
					<td class="esquerda td_Visualizar_favoritos">
			<%
			iIndexSrv = IntQueryString("Servidor", 1)
			iIdioma = IntQueryString("iIdioma", 0)
			iSrvCorrente = cInt(iIndexSrv)
			listaSelecionada = IntQueryString("listaSelecionada", 0)
			%>
				<span><%=getTermo(global_idioma, 8321, "Lista de favoritos", 0)%></span>
					<div class="combo_lista_favoritos" id="combo_lista_favoritos" ></div>
					<span class="div_selecao_favoritos_botoes renomear">
					<%	
						Response.Write "&nbsp;"
						Response.Write "<input type='button' value='"&getTermo(global_idioma, 8322, "Renomear", 0)&"' onclick=javascript:altera_descricao_favoritos("&iIndexSrv&");>"
						Response.Write "&nbsp;&nbsp;"
                        mensagemExclusao = getTermoHtml(global_idioma, 8339, "Confirma a exclusão da lista de favoritos?", 0)
                        Response.Write "<input type='button' value='"&getTermo(global_idioma, 167, "Excluir", 0)&"' onclick=javascript:excluir_favorito("&iIndexSrv&",'"&Server.URLEncode(mensagemExclusao)&"');>"
					 %>
					</span>
					</td>
				</tr>
				<tr>
					<td>
						<div class="ficha_favoritos" id="ficha_favoritos"><span class="span_imagem icon_16 mozilla_blu"></span></div>
					</td>
				</tr>
			</table>
			<script>
			$(document).ready(function () {
				atualiza_Lista_Favoritos(<%=iSrvCorrente %>, <%=listaSelecionada%>);
			});
			</script>
			<% 
		else
			Response.Write "<span><b>"&getTermo(global_idioma, 8347, "Nenhum favorito cadastrado", 0)&"</b></span>"
		end if
	end if
%>	