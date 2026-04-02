 <% if (repositorio_institucional = 1) then %>
    <!-- #include file='acelerador_busca.asp' -->
<% end if %>

<% bDiretorioPadrao = ((sDiretorioArq <> "orcamento") and (sDiretorioArq <> "certidao")) %>

<table class="max_width max_height">
<% if repositorio_institucional = 0 or global_cfg_hab_banner_ri = 1 then %>
<tr>
	<td class="td_center_middle td_home">
		<div id="home_slide_content_div">
   
    <% if (global_numero_serie = 3156) then 'ITA %>
        <span class='span_imagem vazio' style="overflow:hidden; height: 288px; width: 960px"></span>
            <% 
                if (config_multi_servbib = 0) then
            %>
                    <map name="MapBackground" id="MapBackground">
                      <area shape="rect" coords="376,2,428,60" href="http://www.ifi.cta.br/" target="_blank" alt="" />
                      <area shape="rect" coords="221,3,269,61" href="http://www.iae.cta.br/" target="_blank" alt="" />
                      <area shape="rect" coords="541,4,588,59" href="http://www.ieav.cta.br/" target="_blank" alt="" />
                      <area shape="rect" coords="694,4,740,57" href="http://www.ipev.cta.br/" target="_blank" alt="" />
                      <area shape="rect" coords="374,160,585,175" href="http://www.iae.cta.br/" target="_blank" alt="" />
                      <area shape="rect" coords="335,179,622,194" href="http://www.ifi.cta.br/" target="_blank" alt="" />
                      <area shape="rect" coords="382,199,576,214" href="http://www.ieav.cta.br/" target="_blank" alt="" />
                      <area shape="rect" coords="355,219,602,234" href="http://www.ipev.cta.br/" target="_blank" alt="" />
                    </map>
            <% 
                    color = "#FCCC00"
                    top = "-180px"
                    width = "300px"
                else
                    color = "#123666"
                    top = "-45px"
                    width = "150px"
                end if
            if (repositorio_institucional = 0) then
				%>
				<div style="height: 288px; width: 960px; text-align: left">
					<div style="position: relative; margin-top: <%=top%>; width: <%=width%>; text-align: center; font-size: 13px; color: <%=color%>; white-space:nowrap;">
						<b>
							Publicações Eletrônicas
						</b>
					  <br />
						<div style="text-align: left; display:inline-block;width: 100px;">
							<a href="http://www.bibl.ita.br/pubacessorestrito.htm" target="_blank" style="color: <%=color%>;">acesso restrito</a>
							<br />
							<a href="http://www.bibl.ita.br/pubacessopublico.htm" target="_blank" style="color: <%=color%>;">acesso público</a>
						</div>
					</div>
				</div>
		<%	end if
		 end if	

        if (repositorio_institucional) = 0 then
		    if ((config_multi_servbib = 0) and (bDiretorioPadrao = true)) then
			    Set ROService = ROServer.CreateService("Web_Consulta")
			    XMLMensagens = ROService.MensagensPublicas
		
			    TrataErros(1)
		
			    Set ROService = nothing

			    if (XMLMensagens <> "") and (left(XMLMensagens, 5) = "<?xml") then 
				    Set xmlDoc = CreateObject("Microsoft.xmldom")
				    xmlDoc.async = False
				    xmlDoc.loadxml XMLMensagens
				    Set xmlRoot = xmlDoc.documentElement
				    if xmlRoot.nodeName = "MENSAGENS" then
					    QtdMsg = xmlRoot.attributes.getNamedItem("QUANTIDADE").value
				
					    if QtdMsg > 0 then %>

				    <div id="slide_mensagens_div">
					    <div id="slide_mensagens_handle_fundo_div"></div>
							    <a id="slide_mensagens_handle" href="#" title="<%=getTermo(global_idioma, 6994, "Ocultar últimos avisos", 0)%>">
								    <%=getTermo(global_idioma, 6993, "Últimos avisos", 0)%>
						    &nbsp;
						    <span class="span_imagem icon_12 icon-seta-baixo" id="slide_mensagens_seta_img"></span>
					    </a>
					    <div id="slide_mensagens_fundo_div">
					    </div>
					    <div id="slide_mensagens_conteudo_div">
						    <div id="area_mensagens_div">

								    <%
									    For Each xmlMsg In xmlRoot.childNodes
										    '*************************************
										    'MENSAGENS PARA O USUÁRIO
										    '*************************************
										    if xmlMsg.nodeName  = "MENSAGEM" then												
											    MSG = xmlMsg.attributes.getNamedItem("MSG").value
											    TITULO = xmlMsg.attributes.getNamedItem("TITULO").value
											    DATA = xmlMsg.attributes.getNamedItem("DATA").value

											    Response.Write "<div style='padding-left: 25px; padding-right: 25px; text-align: left;'>"
											    Response.Write "<div class='mensagens_titulo'>" & TITULO & "</div>"
											    Response.Write "<div class='mensagens_data'>" & DATA & "</div>"
                                                Response.Write "<div class='mensagem_fonte'>"
											    Response.Write MSG
                                                Response.Write "</div>"
											    Response.Write "</div>"

											    QtdMsg = QtdMsg - 1
											    if (QtdMsg > 0) then
												    Response.Write "<div class=""separador_mensagens_div""></div>"
											    else
												    Response.Write "<div style='padding-bottom: 25px;'></div>"
											    end if
										    end if
									    Next						
								    %>

					    </div>
				    </div>
			    </div>

				    <script type="text/javascript">
					    var mensagensVisiveis = true;

					    $("#slide_mensagens_handle").click(function (e) {

					    
					        if (mensagensVisiveis) {
							    $("#slide_mensagens_div").animate({ marginTop: 268 });
							    $("#slide_mensagens_seta_img").removeClass('icon-seta-baixo');
							    $("#slide_mensagens_seta_img").addClass('icon-seta-cima');
							    $("#slide_mensagens_handle").attr("title", "<%=getTermo(global_idioma, 6995, "Exibir últimos avisos", 0)%>");
						    } else {
							    $("#slide_mensagens_div").animate({ marginTop: 0 });
							    $("#slide_mensagens_seta_img").removeClass('icon-seta-cima');
							    $("#slide_mensagens_seta_img").addClass('icon-seta-baixo');
							    $("#slide_mensagens_handle").attr("title", "<%=getTermo(global_idioma, 6994, "Ocultar últimos avisos", 0)%>");
						    }
						    mensagensVisiveis = !mensagensVisiveis;
						    e.preventDefault();
					    });
				    </script>

					    <%
					    end if
				    end if
			    end if 
		    end if
        end if %>
		</div>
        </td>
</tr>
<% 
end if

if (global_ultimas_aquisicoes = 1 and bDiretorioPadrao = true) then 
	Response.Write "<tr id='ultimas_aquisicoes_tr'>"
	Response.Write "<td id='ultimas_aquisicoes_td' style='background-color: white;'>"
	Response.Write "<div id='conteudo'>"

	if veio_de_index = 1 then 
		if (config_multi_servbib = 1) then 
			Response.Write "<div id='ultimas_aquisicoes_cabecalho_div_select_biblioteca'>"
			Response.Write "<span class='mensagens_titulo'>" &getTermo(global_idioma, 7045, "Últimas aquisições", 0) & "</span>"
			%><!-- #include file="busca_bib_ultimas_aquisicoes.asp"--><%		
		else 
			Response.Write "<div id='ultimas_aquisicoes_cabecalho_div'>"
			if (repositorio_institucional = 1) then
				Response.Write "<span class='mensagens_titulo'>" &getTermo(global_idioma, 9236, "Últimos documentos", 0)& "</span>"
			else
				Response.Write "<span class='mensagens_titulo'>" &getTermo(global_idioma, 7045, "Últimas aquisições", 0)& "</span>"
			end if
		end if 
		Response.Write "</div>"
		Response.Write "<div id='ultimas_aquisicoes_area_conteudo_div'>"
		Response.Write "</div>"
	end if

	Response.Write "</div>"
	Response.Write "</td>"
	Response.Write "</tr>"
end if

if (global_hab_levantamento_bib = 1 and bDiretorioPadrao = true and repositorio_institucional = 0) then

	Response.write "<tr id='levantamento_bibibliografico'  style='background-color: white;'>"
	Response.write "</tr>"

end if

Response.write "</table>"

if (global_ultimas_aquisicoes = 1 and bDiretorioPadrao = true) then %>
	<script type="text/javascript">
	    $(document).ready(function () {
			var indexSrv = 1;
			if(sessionStorage.iIndexSrvAquisicao) {
				indexSrv = sessionStorage.iIndexSrvAquisicao
			}
			
	        atualizaUltimasAquisicoes(indexSrv);
		});
	</script>
<%end if

if (global_hab_levantamento_bib = 1 and bDiretorioPadrao = true and repositorio_institucional = 0) then %>
	<script type="text/javascript">
	    $(document).ready(function () {
			var indexSrv = 1;
			obterLevantamentosBibliografico(indexSrv);
		});
	</script>
<%end if%>