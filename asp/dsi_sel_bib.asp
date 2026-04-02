<% Response.Buffer=True %>
<!DOCTYPE html>
<% 
    htmlClass = ""
    if Request.Cookies("contraste") = "1" then
        htmlClass = "class='contraste'"
    end if
%>
<html <%=htmlClass%>>
<head>
<%
sDiretorioArq = "asp"
%>
<!-- #include file="../config.asp" -->
<!-- #include file="../idiomas/idiomas.asp" -->
<!-- #include file="../libasp/header.asp" -->
<!-- #include file="../libasp/funcoes.asp" -->
<title>::<%=getTermo(global_idioma, 3, "Biblioteca", 0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link href="../inc/estilo_padrao.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_padrao.css" rel="stylesheet" type="text/css" />
<% if config_css_estatico = 1 then %>
	<link href="../inc/estilo.css" rel="stylesheet" type="text/css" /> 
<% else %>
	<link href="../libasp/estilo.asp?idioma=<%=global_idioma%>" rel="stylesheet" type="text/css" />
<% end if %>
<link href="../inc/contraste.css" rel="stylesheet" type="text/css" />
<link href="../inc/imagem_contraste.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../scripts/funcoes.js?b=<%=global_build%>"></script>
<script type="text/javascript">
	function selecao_todos() {
		var desc = "dsi_sel";	
		var doc = document;
		
		var bCheck = true;
		var ckTodos = doc.getElementById('ckTodos');
		if (ckTodos != null) {
			bCheck = ckTodos.checked;
		}
		
		tag = doc.getElementsByTagName('input');
		for(i=tag.length-1;i>=0;i--) {
			if (tag[i].type == "checkbox") {
				if (tag[i].id.substring(0,desc.length) == desc) {
					var codSel = tag[i].id.substring(desc.length,tag[i].id.length);
					tag[i].checked = bCheck;
				}
			}
		}
	}
	
	function pega_selecao() {
		var desc = "dsi_sel";	
		var doc = document;
		var sSel = "";
				
		tag = doc.getElementsByTagName('input');
		for(i=tag.length-1;i>=0;i--) {
			if (tag[i].type == "checkbox") {
				if (tag[i].id.substring(0,desc.length) == desc) {
					var codSel = tag[i].id.substring(desc.length,tag[i].id.length);
					if (tag[i].checked) {
						if (sSel != "") {
							sSel = sSel + ",";
						}
						sSel = sSel + codSel;
					}
				}
			}
		}
		
		var sURL = "dsi_sel_bib.asp?acao=alterar&usuario=<%=Session("codigo_usuario")%>&selecao=" + sSel;
		document.location = sURL;

		return true;
	}
</script>
</head>
<% 
	iIndexSrv = Session("Servidor_Logado")
	
	if iIndexSrv = "" then
		iIndexSrv = 1
	end if

	'O índice iIndexSrv que define em qual servidor será realizada a pesquisa 
	%><!-- #include file="../libasp/updChannelProperty.asp" --><%

	if len(trim(Session("codigo_usuario"))) = 0 then
		Response.Write "<script type='text/javascript'>alert('"&getTermo(global_idioma, 1474, "Sessão expirada", 0)&".\n"&getTermo(global_idioma, 4097, "Favor logar novamente.", 0)&"');parent.fechaPopup();</script>"
		Response.End()
	end if
	
	if request.querystring("acao") = "alterar" then	
		codUsu = Request.QueryString("usuario")
		sSel = Request.QueryString("selecao")
	
		On Error Resume Next
		SET ROService = ROServer.CreateService("Web_Consulta")
		ROService.IncluiDadosDSI codUsu, 1, sSel
		SET ROService = nothing
		TrataErros(1)

		Response.Write "<script type='text/javascript'>parent.document.location='../index.asp?content=dsi&modo_busca=" & GetModo_Busca & "&iBanner="&global_tipo_banner&"&iIdioma=" & global_idioma & "';</script>"
		Response.Write "<script type='text/javascript'>parent.fechaPopup();</script>"	
	else
%>
<body class="popup">
<script type="text/javascript">parent.fechaLoadingPopup();</script>
	<table class="max_width centro">
	<tr>
		<td class="td_center_top">
            <div style="height: 420px; max-height: 420px; overflow: auto;">
				<br />
				<% 
                    '------------------------
                    'Busca registros
                    Set ROService = ROServer.CreateService("Web_Consulta")	
                    sXMLSel = ROService.DadosTabela(roTAB_DSI_BIB, Session("codigo_usuario"), 0,repositorio_institucional)
                    Set ROService = nothing
                    TrataErros(1)
						
					iQtde = 0
						
					if (left(sXMLSel,5) = "<?xml") then
						'Processa o XML
						Set xmlDoc = CreateObject("Microsoft.xmldom")
						xmlDoc.async = False
						xmlDoc.loadxml sXMLSel
						Set xmlRoot = xmlDoc.documentElement
							
						iQtde = CLng(xmlRoot.attributes.getNamedItem("QUANTIDADE").value)
					
						'Verifica se a tabela foi encontrada e se possui registros
						if xmlRoot.attributes.getNamedItem("QUANTIDADE").value <> "0" then
							Response.Write "<table class='tab_borda_selecao' style='display: inline-table'><tr><td>"
								
							Response.Write "<table class='max_width' style='border-spacing: 1px; padding: 0'>"

							Response.Write "<tr style='height: 20px'>"

							Response.Write "<td class='td_tabelas_titulo centro' style='width: 35px'>"
							Response.Write "<input type='checkbox' id='ckTodos' value='0' onclick='selecao_todos();'/>"
							Response.Write "</td>"
								
							Response.Write "<td class='td_tabelas_titulo centro' style='width: 270px'>"
							Response.Write getTermo(global_idioma, 25, "Descrição", 0)
							Response.Write "</td>"
								
							Response.Write "</tr>"
								
							iSeq = 1
								
							For Each xmlTabela In xmlRoot.childNodes
								if xmlTabela.nodeName  = "REGISTRO" then
									if (iSeq mod 2) > 0 then '### IMPAR
										td_class = "td_tabelas_valor2"
									else '### PAR
										td_class = "td_tabelas_valor1"
									end if									
									
									sDesc  = xmlTabela.attributes.getNamedItem("DESCRICAO").value
									sCod   = xmlTabela.attributes.getNamedItem("CODIGO").value
									sCheck = xmlTabela.attributes.getNamedItem("SELECIONADO").value
										
									if (sCheck = "1") then
										sCheck = "checked"
									else
										sCheck = ""
									end if
										
									Response.Write "<tr style='height: 20px'>"
										
									Response.Write "<td class='"&td_class&" centro'>"
									Response.Write "<input type='checkbox' id='dsi_sel" & sCod & "' value='" & sCod & "' " & sCheck & "/>"
									Response.Write "</td>"
										
									Response.Write "<td class='"&td_class&" esquerda'>&nbsp;"
									Response.Write sDesc
									Response.Write "</td>"
									
									Response.Write "</tr>"
										
									iSeq = iSeq + 1
								end if
							Next
								
							Response.Write "</table>"
							Response.Write "</td></tr></table>"
						end if
							
						Set xmlRoot = nothing
						Set xmlDoc = nothing
					end if
                %> 
            </div>
		</td>
    </tr>
    <tr style="height:30px">
        <td>
            <p class="centro">
			<input type='button' value='<%=getTermo(global_idioma, 510, "Alterar", 0)%>' onclick='return pega_selecao();' />
			&nbsp;&nbsp;
            <input type='button' value='<%=getTermo(global_idioma, 5, "Cancelar", 0)%>' onclick='parent.fechaPopup()' />
            </p>
        </td>            
	</tr>
	</table>
</body>
<% end if %>
</html>