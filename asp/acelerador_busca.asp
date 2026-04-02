
<script type="text/javascript" src="scripts/acelerador_busca.js"></script>

<% 

contador = 0
if (global_cfg_hab_tese_dissertacao_ri = 1) then
    contador = contador + 1
end if
if (global_cfg_hab_busca_material_ri = 1) then
    contador = contador + 1
end if
if (global_cfg_hab_busca_repositorio_ri = 1) then
    contador = contador + 1
end if

Select case contador
	case 1
		classeAcelerador = "panelbar-acelerador-uma-coluna" 
	case 2
		classeAcelerador = "panelbar-acelerador-duas-colunas"
	case 3
		classeAcelerador = "panelbar-acelerador-tres-colunas"
End Select

if (contador > 0) then
%>
<div id="div-aceleradores-busca">
  <% if (global_cfg_hab_tese_dissertacao_ri = 1) then %>
      <div class="div-acelerador-busca <%=classeAcelerador%>" id="painelTeseDissertacao">
          <div class="titulo-acelerador-busca linha">
              <span>
                  <%=getTermo(global_idioma, 2194, "Teses", 0) & "/" & getTermo(global_idioma, 9396, "Dissertações", 0)%>
              </span>
          </div>
          <div class="corpo-acelerador-busca">
            <div id="conteudo-tese-dissertacao">
                <%=getTermo(global_idioma, 9828, "Visualizar teses e dissertações agrupadas por programa e área de concentração.", 0)%>
                <span class="span_imagem div_imagem_right_3 icon_9 mais-png"></span>
                <a class='link_classic2' href="javascript:LinkTeseDissertacao();">
                    <%=getTermo(global_idioma, 9273, "Ver todos", 0)%>
                </a>
            </div>
          </div>  
      </div>
  <% end if

  if (global_cfg_hab_busca_material_ri = 1) then %>
      <div class="div-acelerador-busca <%=classeAcelerador%>" id="painel_material">
      </div>
  <% end if

  if (global_cfg_hab_busca_repositorio_ri = 1) then %>
      <div class="div-acelerador-busca <%=classeAcelerador%>" id="painel_repositorio">
      </div>
  <% end if %>
</div>
<% end if %>