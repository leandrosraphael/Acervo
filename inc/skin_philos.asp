<%
'//#########################################################################//
'//             AS IMAGENS DEVEM ESTAR NO DIRETÓRIO IMAGENS                 //
'//#########################################################################//
'---------------------------------------------------------------------------//
'BANNER - IMAGEM NO TOPO DO SITE (tamanho ideal: 970 x  72 pixels)
if (global_idioma = 1) OR (global_idioma = 3) then
	css_imagem_banner = "banner_ph_esp.png"
elseif (global_idioma = 2) then
	css_imagem_banner = "banner_ph_ing.png"
elseif (global_idioma = 5) then
	css_imagem_banner = "banner_ph_gal.png"
elseif (global_idioma = 6) then
	css_imagem_banner = "banner_ph_eus.png"
else
	css_imagem_banner = "banner_ph.png"
end if
'---------------------------------------------------------------------------//
'ALINHAMENTO DO BANNER
'OPÇÔES VÁLIDAS: left,center ou right
css_alinhamento_banner = "center"
'---------------------------------------------------------------------------//
'COR DE FUNDO DO BANNER
css_background_banner = "#FFFFFF"      
'---------------------------------------------------------------------------//
'COR DA FONTE DA FRASE APARECENDO SOBRE O BANNER 
css_banner_cor_fonte = "#0070C0"
'---------------------------------------------------------------------------//
'TAMANHO DA FONTE DA FRASE APARECENDO SOBRE O BANNER 
css_banner_tamanho_fonte = "28"
'---------------------------------------------------------------------------//
'ALINHAMENTO DA FRASE APARECENDO SOBRE O BANNER 
'OPÇÔES VÁLIDAS: left,center ou right
css_alinhamento_frase_banner = "center"
'---------------------------------------------------------------------------//
'IMAGEM NA PÁGINA DE ENTRADA (tamanho ideal: 970 x 290 pixels)
if (repositorio_institucional = 0) then
	if (global_idioma = 1) OR (global_idioma = 3) OR (global_idioma = 5) OR (global_idioma = 6) then
		css_imagem_home = "background_ph-es.jpg"
	elseif (global_idioma = 2) then
		css_imagem_home = "background_ph-en.jpg"
	else
		css_imagem_home = "background_ph.jpg"
	end if
else
	css_imagem_home = "background-ri.jpg"
end if

'---------------------------------------------------------------------------//

'//#########################################################################//
'//                             CUSTOMIZAÇÃO                                //
'//#########################################################################//
	
'//------------------------CONFIGURAÇÕES GERAIS-----------------------------//

css_geral_fonte = """Segoe UI"", Frutiger, ""Frutiger Linotype"", ""Dejavu Sans"", ""Helvetica Neue"", Arial, sans-serif"
css_geral_tamanho_fonte = "13"
css_geral_background_externo = "#F2F2F2"
css_geral_botao_background = "#FFFFFF"
css_geral_botao_tamanho_fonte = "13"
css_geral_botao_cor_fonte = "#000000"
css_geral_botao_cor_margem = "#aaaaaa"
css_geral_link_servicos_normal = "#000000"
css_geral_link_servicos_over = "#009999"
css_geral_background = "#FFFFFF"
css_geral_icones_cor = "#0070c0"
'//-------------------CONFIGURAÇÃO DO MENU PRINCIPAL------------------------//

css_menu_background = "#0070C0"
css_menu_link_normal = "#FFFFFF"
css_menu_link_over = "#00FFFF"
css_menu_link_font_weight = "400"

'//----------------------CONFIGURAÇÃO DOS SERVIÇOS--------------------------//

css_servicos_menu_link_normal = "#000000"
css_servicos_menu_link_over = "#009999"
css_mensagens_background = "#FFFFFF"


'//----------------------CONFIGURAÇÃO DAS BUSCAS----------------------------//

css_busca_ativa = "#DDDDDD"
css_busca_inativa = "#BABABA"
css_busca_tamanho_fonte = "13"
css_busca_case_fonte = "Minusculo"
css_abas_link_normal = "#000000"
css_abas_link_over = "#FFFFFF"
css_busca_inputs = "#006699"

'//----------------------CONFIGURAÇÃO DA PAGINAÇÃO--------------------------//

css_paginacao_background = "#DEDEDE"
css_paginacao_link_normal = "#000000"
css_paginacao_link_over = "#006699"
css_paginacao_link_atual = "#FF0000"

'//-----------------------CONFIGURAÇÃO DOS GRIDS----------------------------//

css_grid_cabecalho = "#DDDDDD"
css_grid_cabecalho_fonte = "#000000"

'//----------------------CONFIGURAÇÃO DOS POP UPS---------------------------//

css_popup_cabecalho = "#0070c0"
css_popup_cabecalho_cor_fonte = "#FFFFFF"

'//----------------------CONFIGURAÇÃO DA FICHA DE DETALHES------------------//

css_detalhe_esquerda = "#DDDDDD"
css_detalhe_direita = "#FFFFFF"

'//-------------------CONFIGURAÇÃO DO DESTAQUE DAS PALAVRAS-----------------//

css_destaca_palavras = "#FFFF00"

'//#########################################################################//

%>