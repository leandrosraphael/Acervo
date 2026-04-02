
<%	
on error resume next
Set ROService = ROServer.CreateService("Web_Consulta")
XMLConfigGlobais = ROService.ConfigGlobaisXML(Retorna_IP)

Set ROService = nothing
	
if len(trim(XMLConfigGlobais)) <> 0 then
	if (left(XMLConfigGlobais,5) = "<?xml") then
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml XMLConfigGlobais
		Set xmlConfiguracao = xmlDoc.documentElement
		'Dados da licença
		global_versao 		= getValorConfiguracaoInt(xmlConfiguracao, "DTA_VERSAO")
		global_dsi 			= getValorConfiguracaoInt(xmlConfiguracao, "DTA_DSI")
		global_multibib 	= getValorConfiguracaoInt(xmlConfiguracao, "DTA_MULTIBIB")
		global_marc 		= getValorConfiguracaoInt(xmlConfiguracao, "DTA_MARC")
		global_legislacao 	= getValorConfiguracaoInt(xmlConfiguracao, "DTA_JURIDICO")
		global_espanha 	    = getValorConfiguracaoInt(xmlConfiguracao, "DTA_ESPANHA")
		global_mobile 	    = getValorConfiguracaoInt(xmlConfiguracao, "DTA_MOBILE")
        global_bibdigital   = getValorConfiguracaoInt(xmlConfiguracao, "DTA_BIBDIGITAL")
	
		global_academico 	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_ACADEMICO")

		'Dados da configuração local
		global_IP_Local 		= getValorConfiguracaoInt(xmlConfiguracao, "CFG_IP_LOCAL")
		global_bib_dif_usu 		= getValorConfiguracao(xmlConfiguracao, "CFG_RESERVA_BIBDIFUSU")
		global_bib_indica 		= getValorConfiguracaoInt(xmlConfiguracao, "CFG_RESERVA_INDICABIB")
		global_renovacao 		= getValorConfiguracaoInt(xmlConfiguracao, "CFG_RENOVA")
		global_Usa_Rm_Terminal 	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_USARM")
		Bloq_Sug_Aquis 			= getValorConfiguracao(xmlConfiguracao, "CFG_SUGAQ")
	
		'Dados da configuração geral
		global_hab_servicos 		        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_SERVICOS")
		config_habilita_autoridades         = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_CAT_AUTORIDADES")
		global_hab_bus_leg 			        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_BUS_LEG")
		global_hab_troca_senha 		        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_TROCA_SENHA")
		global_hab_envio_email_ms 	        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_ENVIO_EMAIL_MS")
		global_banner 				        = getValorConfiguracao(xmlConfiguracao, "CFG_BANNER")
		global_midia_hab_cgi 		        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_MIDIA_HAB_CGI")
		global_midia_caminho_virtual        = getValorConfiguracao(xmlConfiguracao, "CFG_MIDIA_CAMINHO_VIRTUAL")
		global_midia_caminho_fisico         = getValorConfiguracao(xmlConfiguracao, "CFG_MIDIA_CAMINHO_FISICO")
		global_num_linhas			        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_NUM_RESULTADOS")
		config_css_estatico			        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_CSS_TIPO")
		global_portal_acad                  = getValorConfiguracaoInt(xmlConfiguracao, "CFG_PORTALACAD")
		global_portal_acad_url              = getValorConfiguracao(xmlConfiguracao, "CFG_PORTALACAD_URL")
		global_exibe_capa                   = getValorConfiguracaoInt(xmlConfiguracao, "CFG_EXIBE_CAPA")
		global_email_lembranca_senha        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_LEMBRANCA_SENHA")
		global_bib_curso                    = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_BIB_CURSO")
		global_avaliacao_online	            = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_AVALIACAO_ONLINE")
		global_avaliacao_autenticada        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_AVALIACAO_ONLINE_AUTENTICADA")
		global_salvar_marc                  = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_SALVAR_MARC")
		global_busca_descricao_complementar = getValorConfiguracaoInt(xmlConfiguracao, "CFG_BUSCA_DESCRICAO_COMPLEMENTAR")
		global_eds	     		            = getValorConfiguracaoInt(xmlConfiguracao, "CFG_BUSCA_EDS")
		global_nome_instituicao				= getValorConfiguracao(xmlConfiguracao, "CFG_NOME_INSTITUICAO")
		global_ultimas_aquisicoes			= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_ULTIMAS_AQUISICOES")
		global_dias_ultimas_aquisicoes		= getValorConfiguracaoInt(xmlConfiguracao, "CFG_DIAS_ULTIMAS_AQUISICOES")
		global_facebook_curtir				= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_FACEBOOK_CURTIR")
		global_facebook_appid				= getValorConfiguracao(xmlConfiguracao, "CFG_FACEBOOK_APPID")
		global_twitter						= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_TWITTER")
		global_inf_pessoais                 = getValorConfiguracaoInt(xmlConfiguracao, "CFG_DADOS_PESSOAIS")
		global_arquivo						= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_ARQUIVO")
		global_hab_bus_subloc               = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_BUS_SUBLOC")
		global_desc_subloc		            = getValorConfiguracao(xmlConfiguracao, "CFG_DESC_SUBLOC")
		global_emprestimo_digital           = getValorConfiguracaoInt(xmlConfiguracao, "CFG_EMPRESTIMODIGITAL")
		global_busca_facetada               = getValorConfiguracaoInt(xmlConfiguracao, "CFG_EXIBE_BUSCA_FACETADA")
		global_atualizar_email              = getValorConfiguracaoInt(xmlConfiguracao, "CFG_ATUALIZAR_EMAIL")
		global_hab_sobre_mobile             = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_SOBRE_MOBILE")
		global_titulo_sobre_mobile          = getValorConfiguracao(xmlConfiguracao, "CFG_TITULO_SOBRE_MOBILE")

		'Dados da configuração de idiomas
		global_idioma			= getValorConfiguracaoInt(xmlConfiguracao, "CFG_IDIOMA_PADRAO")
		global_idioma_portugues	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_IDIOMA_PORTUGUES")
		global_idioma_espanhol	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_IDIOMA_ESPANHOL")
		global_idioma_ingles	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_IDIOMA_INGLES")
		global_idioma_catalao	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_IDIOMA_CATALAO")
		global_idioma_alemao	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_IDIOMA_ALEMAO")
		global_idioma_galego	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_IDIOMA_GALEGO")
		global_idioma_euskera	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_IDIOMA_EUSKERA")
	
		'Configuração do Streaming
		global_hab_streaming			= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_STREAMING")
		global_streaming_caminho_fisico	= getValorConfiguracao(xmlConfiguracao, "CFG_STREAMING_CAMINHO_FISICO")
		global_streaming_caminho_virtual= getValorConfiguracao(xmlConfiguracao, "CFG_STREAMING_CAMINHO_VIRTUAL")
	
		global_recursos_avancados	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_RECURSOS_AVANCADOS")

		global_hab_google_maps		= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_GOOGLE_MAPS")
		global_chave_google_maps	= getValorConfiguracao(xmlConfiguracao, "CFG_CHAVE_GOOGLE_MAPS")

		'Controle de limpeza do cache de imagens das últimas aquisições
		LimparCacheImagensUltimasAquisicoes = getValorConfiguracaoInt(xmlConfiguracao, "LIMPAR_CACHE_ULTIMAS_AQUISICOES")

		global_pesquisar_fonte_avulsa	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_PESQUISAR_FONTE_AVULSA")

		global_hab_levantamento_bib	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_LEVANTAMENTO_BIB")

		global_hab_bus_bib						= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_BUS_BIB")
		global_url_ajuda_customizada			= getValorConfiguracao(xmlConfiguracao, "URL_AJUDA_CUSTOMIZADA")
		global_cfg_exemplar_tombo				= getValorConfiguracaoInt(xmlConfiguracao, "CFG_EXEMPLAR_TOMBO")
		global_cfg_utilizar_rybena				= getValorConfiguracaoInt(xmlConfiguracao, "CFG_UTILIZAR_RYBENA")

		global_cfg_integrar_editora_fgv			= getValorConfiguracaoInt(xmlConfiguracao, "CFG_INTEGRAR_EDITORA_FGV")
		global_cfg_exigir_login_terminal_web	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_EXIGIR_LOGIN_TERMINAL_WEB")

		global_cfg_terminal_rep_institucional	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_TERMINAL_REP_INSTITUCIONAL")
		global_cfg_nome_rep_institucional		= getValorConfiguracao(xmlConfiguracao, "CFG_NOME_REP_INSTITUCIONAL")
		global_cfg_url_rep_institucional		= getValorConfiguracao(xmlConfiguracao, "CFG_URL_REP_INSTITUCIONAL")
		global_cfg_nome_menu_repositorio_inst	= getValorConfiguracao(xmlConfiguracao, "CFG_NOME_MENU_REPOSITORIO_INST")

		global_cadastro_usuario_externo					= getValorConfiguracaoInt(xmlConfiguracao, "CFG_CADASTRO_USUARIO_EXTERNO")
		global_cfg_cadastrar_empresa_usuario_externo	= getValorConfiguracaoInt(xmlConfiguracao, "CFG_CADASTRAR_EMPRESA_USUARIO_EXTERNO")
		global_cfg_hab_tese_dissertacao_ri				= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_TESE_DISSERTACAO_RI")
		global_cfg_hab_indicadores_ri					= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_INDICADORES_RI")
		global_cfg_hab_grafico_tes_ano_pro_ri			= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_GRAFICO_TES_ANO_PRO_RI")
		global_cfg_hab_grafico_tes_pro_ano_ri			= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_GRAFICO_TES_PRO_ANO_RI")
		global_cfg_hab_grafico_tes_mat_ano_ri			= getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_GRAFICO_TES_MAT_ANO_RI")
        global_cfg_hab_banner_ri                        = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_BANNER_RI")
        global_cfg_hab_busca_material_ri                = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_BUSCA_MATERIAL_RI")
        global_cfg_hab_busca_repositorio_ri             = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_BUSCA_REPOSITORIO_RI")
        global_cfg_hab_grafico_mat_ano_lin_ri           = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_GRAFICO_MAT_ANO_LIN_RI")
        global_cfg_hab_grafico_mat_ano_bar_ri           = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_GRAFICO_MAT_ANO_BAR_RI")

		if (repositorio_institucional = 1) then
			global_cfg_habilitar_links_uteis = 0
			global_cfg_nome_links_uteis = ""
		else
			global_cfg_habilitar_links_uteis = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HABILITAR_LINKS_UTEIS")
			global_cfg_nome_links_uteis = getValorConfiguracao(xmlConfiguracao, "CFG_NOME_LINKS_UTEIS")
		end if

        global_cfg_hab_emprestimo_app = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HABILITAR_EMPRESTIMO_APP")
        global_cfg_habilitar_certidao_negativa = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_CERTIDAO_NEGATIVA")
        global_cfg_usa_repos_digital = getValorConfiguracaoInt(xmlConfiguracao, "CFG_USA_REPOS_DIGITAL")
		global_cfg_hab_qrcode_exemplar = getValorConfiguracaoInt(xmlConfiguracao, "CFG_HAB_QRCODE_EXEMPLAR")

		Set xmlDoc = nothing
		Set xmlConfiguracao = nothing
	end if
end if
%>