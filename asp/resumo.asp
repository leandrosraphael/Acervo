<!-- #include file="../libasp/funcoes.asp" -->
<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->
<%
if (sMsgErro <> "") then
	Response.Write sMsgErro
else
	if (Request.QueryString("veio_de") = "voltar") then
	%>
		<!-- #include file="../asp/ler_parametros_busca.asp" -->
	<%
		sURL_VOLTAR = ""
	else
		sDados = trim(Request.Form("dados"))
		if (trim(Request.Form("dados_outro")) <> "") then
			sDados = sDados & " " & trim(Request.Form("dados_outro"))
		end if
		
		if (trim(Request.Form("objeto")) <> "") then
			iObjeto = CInt(Request.Form("objeto"))
		else
			iObjeto = 0
		end if
		
		if (trim(Request.Form("contexto")) <> "") then
			iContexto = CInt(Request.Form("contexto"))
		else
			iContexto = 0
		end if
		
		if (trim(Request.Form("imagem")) <> "") then
			iImagem = CInt(Request.Form("imagem"))
		else
			iImagem = 0
		end if
		
		
		if (trim(Request.Form("tipo_campo1")) <> "") then
			sTipoCampo1 = trim(Request.Form("tipo_campo1"))
			if (trim(Request.Form("tipo_campo1")) = "ALFA") then
				sCampo1 = trim(Request.Form("campo1"))
			elseif (trim(Request.Form("tipo_campo1")) = "LOGICO") then
				sCampo1 = trim(Request.Form("campo1"))
			elseif (trim(Request.Form("tipo_campo1")) = "NUM") then
				if (trim(Request.Form("cb_campo1")) = "3") then
					sCampo1 = trim(Request.Form("cb_campo1")) & "|" & _
					          trim(Request.Form("campo1_ini")) & "|" & _
					          trim(Request.Form("campo1_fim"))
				else
					if trim(Request.Form("campo1_ini")) <> "" then
						sCampo1 = trim(Request.Form("cb_campo1")) & "|" & _
								  trim(Request.Form("campo1_ini")) & "|" & _
								  trim(Request.Form("campo1_ini"))
					else
						sCampo1 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo1")) = "DATA") then
				if (trim(Request.Form("cb_campo1")) = "3") then
					if (trim(Request.Form("campo1_ini_d")) <> "") and (trim(Request.Form("campo1_ini_m")) <> "") and _
					   (trim(Request.Form("campo1_ini_a")) <> "") and (trim(Request.Form("campo1_fim_d")) <> "") and _
                       (trim(Request.Form("campo1_fim_m")) <> "") and (trim(Request.Form("campo1_fim_a")) <> "") then
                        sCampo1 = trim(Request.Form("cb_campo1")) & "|" & _
					          trim(Request.Form("campo1_ini_d")) & "/" & _
					          trim(Request.Form("campo1_ini_m")) & "/" & _
					          trim(Request.Form("campo1_ini_a")) & "|" & _
					          trim(Request.Form("campo1_fim_d")) & "/" & _
					          trim(Request.Form("campo1_fim_m")) & "/" & _
					          trim(Request.Form("campo1_fim_a"))
                    else
						sCampo1 = ""
					end if
				else
					if (trim(Request.Form("campo1_ini_d")) <> "") and (trim(Request.Form("campo1_ini_m")) <> "") and _
					   (trim(Request.Form("campo1_ini_a")) <> "") then
						sCampo1 = trim(Request.Form("cb_campo1")) & "|" & _
								  trim(Request.Form("campo1_ini_d")) & "/" & _
								  trim(Request.Form("campo1_ini_m")) & "/" & _
								  trim(Request.Form("campo1_ini_a")) & "|" & _
								  trim(Request.Form("campo1_ini_d")) & "/" & _
								  trim(Request.Form("campo1_ini_m")) & "/" & _
								  trim(Request.Form("campo1_ini_a"))
					else
						sCampo1 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo1")) = "TABELA") then
				if (trim(Request.Form("cod_tabela1")) <> "") and (trim(Request.Form("cb_campo1")) <> "") and _
				   (trim(Request.Form("cb_campo1")) <> "0") then				
					sCampo1 = trim(Request.Form("cod_tabela1")) & "," & trim(Request.Form("cb_campo1"))
				end if
			else
				sCampo1 = ""
			end if
		else
			if (trim(Request.Form("campo1")) <> "") then
				sCampo1 = trim(Request.Form("campo1"))
			else
				sCampo1 = ""
			end if
		end if
		
		if (trim(Request.Form("tipo_campo2")) <> "") then
			sTipoCampo2 = trim(Request.Form("tipo_campo2"))
			if (trim(Request.Form("tipo_campo2")) = "ALFA") then
				sCampo2 = trim(Request.Form("campo2"))
			elseif (trim(Request.Form("tipo_campo2")) = "LOGICO") then
				sCampo2 = trim(Request.Form("campo2"))
			elseif (trim(Request.Form("tipo_campo2")) = "NUM") then
				if (trim(Request.Form("cb_campo2")) = "3") then
					sCampo2 = trim(Request.Form("cb_campo2")) & "|" & _
					          trim(Request.Form("campo2_ini")) & "|" & _
					          trim(Request.Form("campo2_fim"))
				else
					if trim(Request.Form("campo2_ini")) <> "" then
						sCampo2 = trim(Request.Form("cb_campo2")) & "|" & _
								  trim(Request.Form("campo2_ini")) & "|" & _
								  trim(Request.Form("campo2_ini"))
					else
						sCampo2 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo2")) = "DATA") then
				if (trim(Request.Form("cb_campo2")) = "3") then
					if (trim(Request.Form("campo2_ini_d")) <> "") and (trim(Request.Form("campo2_ini_m")) <> "") and _
					   (trim(Request.Form("campo2_ini_a")) <> "") and (trim(Request.Form("campo2_fim_d")) <> "") and _
                       (trim(Request.Form("campo2_fim_m")) <> "") and (trim(Request.Form("campo2_fim_a")) <> "") then
                        sCampo2 = trim(Request.Form("cb_campo2")) & "|" & _
					          trim(Request.Form("campo2_ini_d")) & "/" & _
					          trim(Request.Form("campo2_ini_m")) & "/" & _
					          trim(Request.Form("campo2_ini_a")) & "|" & _
					          trim(Request.Form("campo2_fim_d")) & "/" & _
					          trim(Request.Form("campo2_fim_m")) & "/" & _
					          trim(Request.Form("campo2_fim_a"))
                    else
						sCampo2 = ""
					end if
				else
					if (trim(Request.Form("campo2_ini_d")) <> "") and (trim(Request.Form("campo2_ini_m")) <> "") and _
					   (trim(Request.Form("campo2_ini_a")) <> "") then
						sCampo2 = trim(Request.Form("cb_campo2")) & "|" & _
								  trim(Request.Form("campo2_ini_d")) & "/" & _
								  trim(Request.Form("campo2_ini_m")) & "/" & _
								  trim(Request.Form("campo2_ini_a")) & "|" & _
								  trim(Request.Form("campo2_ini_d")) & "/" & _
								  trim(Request.Form("campo2_ini_m")) & "/" & _
								  trim(Request.Form("campo2_ini_a"))
					else
						sCampo2 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo2")) = "TABELA") then
				if (trim(Request.Form("cod_tabela2")) <> "") and (trim(Request.Form("cb_campo2")) <> "") and _
				   (trim(Request.Form("cb_campo2")) <> "0") then				
					sCampo2 = trim(Request.Form("cod_tabela2")) & "," & trim(Request.Form("cb_campo2"))
				end if
	 		else
				sCampo2 = ""
			end if
		else
			if (trim(Request.Form("campo2")) <> "") then
				sCampo2 = trim(Request.Form("campo2"))
			else
				sCampo2 = ""
			end if
		end if
		
		if (trim(Request.Form("tipo_campo3")) <> "") then
			sTipoCampo3 = trim(Request.Form("tipo_campo3"))
			if (trim(Request.Form("tipo_campo3")) = "ALFA") then
				sCampo3 = trim(Request.Form("campo3"))
			elseif (trim(Request.Form("tipo_campo3")) = "LOGICO") then
				sCampo3 = trim(Request.Form("campo3"))
			elseif (trim(Request.Form("tipo_campo3")) = "NUM") then
				if (trim(Request.Form("cb_campo3")) = "3") then
					sCampo3 = trim(Request.Form("cb_campo3")) & "|" & _
					          trim(Request.Form("campo3_ini")) & "|" & _
					          trim(Request.Form("campo3_fim"))
				else
					if trim(Request.Form("campo3_ini")) <> "" then
						sCampo3 = trim(Request.Form("cb_campo3")) & "|" & _
								  trim(Request.Form("campo3_ini")) & "|" & _
								  trim(Request.Form("campo3_ini"))
					else
						sCampo3 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo3")) = "DATA") then
				if (trim(Request.Form("cb_campo3")) = "3") then
					if (trim(Request.Form("campo3_ini_d")) <> "") and (trim(Request.Form("campo3_ini_m")) <> "") and _
					   (trim(Request.Form("campo3_ini_a")) <> "") and (trim(Request.Form("campo3_fim_d")) <> "") and _
                       (trim(Request.Form("campo3_fim_m")) <> "") and (trim(Request.Form("campo3_fim_a")) <> "") then
                        sCampo3 = trim(Request.Form("cb_campo3")) & "|" & _
					          trim(Request.Form("campo3_ini_d")) & "/" & _
					          trim(Request.Form("campo3_ini_m")) & "/" & _
					          trim(Request.Form("campo3_ini_a")) & "|" & _
					          trim(Request.Form("campo3_fim_d")) & "/" & _
					          trim(Request.Form("campo3_fim_m")) & "/" & _
					          trim(Request.Form("campo3_fim_a"))
                    else
						sCampo3 = ""
					end if
				else
					if (trim(Request.Form("campo3_ini_d")) <> "") and (trim(Request.Form("campo3_ini_m")) <> "") and _
					   (trim(Request.Form("campo3_ini_a")) <> "") then
						sCampo3 = trim(Request.Form("cb_campo3")) & "|" & _
								  trim(Request.Form("campo3_ini_d")) & "/" & _
								  trim(Request.Form("campo3_ini_m")) & "/" & _
								  trim(Request.Form("campo3_ini_a")) & "|" & _
								  trim(Request.Form("campo3_ini_d")) & "/" & _
								  trim(Request.Form("campo3_ini_m")) & "/" & _
								  trim(Request.Form("campo3_ini_a"))
					else
						sCampo3 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo3")) = "TABELA") then
				if (trim(Request.Form("cod_tabela3")) <> "") and (trim(Request.Form("cb_campo3")) <> "") and _
				   (trim(Request.Form("cb_campo3")) <> "0") then
					sCampo3 = trim(Request.Form("cod_tabela3")) & "," & trim(Request.Form("cb_campo3"))
				end if
			else
				sCampo3 = ""
			end if
		else
			if (trim(Request.Form("campo3")) <> "") then
				sCampo3 = trim(Request.Form("campo3"))
			else
				sCampo3 = ""
			end if
		end if
		
		if (trim(Request.Form("tipo_campo4")) <> "") then
			sTipoCampo4 = trim(Request.Form("tipo_campo4"))
			if (trim(Request.Form("tipo_campo4")) = "ALFA") then
				sCampo4 = trim(Request.Form("campo4"))
			elseif (trim(Request.Form("tipo_campo4")) = "LOGICO") then
				sCampo4 = trim(Request.Form("campo4"))
			elseif (trim(Request.Form("tipo_campo4")) = "NUM") then
				if (trim(Request.Form("cb_campo4")) = "3") then
					sCampo4 = trim(Request.Form("cb_campo4")) & "|" & _
					          trim(Request.Form("campo4_ini")) & "|" & _
					          trim(Request.Form("campo4_fim"))
				else
					if trim(Request.Form("campo4_ini")) <> "" then
						sCampo4 = trim(Request.Form("cb_campo4")) & "|" & _
								  trim(Request.Form("campo4_ini")) & "|" & _
								  trim(Request.Form("campo4_ini"))
					else
						sCampo4 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo4")) = "DATA") then
				if (trim(Request.Form("cb_campo4")) = "3") then
					if (trim(Request.Form("campo4_ini_d")) <> "") and (trim(Request.Form("campo4_ini_m")) <> "") and _
					   (trim(Request.Form("campo4_ini_a")) <> "") and (trim(Request.Form("campo4_fim_d")) <> "") and _
                       (trim(Request.Form("campo4_fim_m")) <> "") and (trim(Request.Form("campo4_fim_a")) <> "") then
                        sCampo4 = trim(Request.Form("cb_campo4")) & "|" & _
					          trim(Request.Form("campo4_ini_d")) & "/" & _
					          trim(Request.Form("campo4_ini_m")) & "/" & _
					          trim(Request.Form("campo4_ini_a")) & "|" & _
					          trim(Request.Form("campo4_fim_d")) & "/" & _
					          trim(Request.Form("campo4_fim_m")) & "/" & _
					          trim(Request.Form("campo4_fim_a"))
                    else
						sCampo4 = ""
					end if
				else
					if (trim(Request.Form("campo4_ini_d")) <> "") and (trim(Request.Form("campo4_ini_m")) <> "") and _
					   (trim(Request.Form("campo4_ini_a")) <> "") then
						sCampo4 = trim(Request.Form("cb_campo4")) & "|" & _
								  trim(Request.Form("campo4_ini_d")) & "/" & _
								  trim(Request.Form("campo4_ini_m")) & "/" & _
								  trim(Request.Form("campo4_ini_a")) & "|" & _
								  trim(Request.Form("campo4_ini_d")) & "/" & _
								  trim(Request.Form("campo4_ini_m")) & "/" & _
								  trim(Request.Form("campo4_ini_a"))
					else
						sCampo4 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo4")) = "TABELA") then
				if (trim(Request.Form("cod_tabela4")) <> "") and (trim(Request.Form("cb_campo4")) <> "") and _ 
				   (trim(Request.Form("cb_campo4")) <> "0") then
					sCampo4 = trim(Request.Form("cod_tabela4")) & "," & trim(Request.Form("cb_campo4"))
				end if
			else
				sCampo4 = ""
			end if
		else
			if (trim(Request.Form("campo4")) <> "") then
				sCampo4 = trim(Request.Form("campo4"))
			else
				sCampo4 = ""
			end if
		end if
		
		if (trim(Request.Form("tipo_campo5")) <> "") then
			sTipoCampo5 = trim(Request.Form("tipo_campo5"))
			if (trim(Request.Form("tipo_campo5")) = "ALFA") then
				sCampo5 = trim(Request.Form("campo5"))
			elseif (trim(Request.Form("tipo_campo5")) = "LOGICO") then
				sCampo5 = trim(Request.Form("campo5"))
			elseif (trim(Request.Form("tipo_campo5")) = "NUM") then
				if (trim(Request.Form("cb_campo5")) = "3") then
					sCampo5 = trim(Request.Form("cb_campo5")) & "|" & _
					          trim(Request.Form("campo5_ini")) & "|" & _
					          trim(Request.Form("campo5_fim"))
				else
					if trim(Request.Form("campo5_ini")) <> "" then
						sCampo5 = trim(Request.Form("cb_campo5")) & "|" & _
								  trim(Request.Form("campo5_ini")) & "|" & _
								  trim(Request.Form("campo5_ini"))
					else
						sCampo5 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo5")) = "DATA") then
				if (trim(Request.Form("cb_campo5")) = "3") then
					if (trim(Request.Form("campo5_ini_d")) <> "") and (trim(Request.Form("campo5_ini_m")) <> "") and _
					   (trim(Request.Form("campo5_ini_a")) <> "") and (trim(Request.Form("campo5_fim_d")) <> "") and _
                       (trim(Request.Form("campo5_fim_m")) <> "") and (trim(Request.Form("campo5_fim_a")) <> "") then
                        sCampo5 = trim(Request.Form("cb_campo5")) & "|" & _
					          trim(Request.Form("campo5_ini_d")) & "/" & _
					          trim(Request.Form("campo5_ini_m")) & "/" & _
					          trim(Request.Form("campo5_ini_a")) & "|" & _
					          trim(Request.Form("campo5_fim_d")) & "/" & _
					          trim(Request.Form("campo5_fim_m")) & "/" & _
					          trim(Request.Form("campo5_fim_a"))
                    else
						sCampo5 = ""
					end if
				else
					if (trim(Request.Form("campo5_ini_d")) <> "") and (trim(Request.Form("campo5_ini_m")) <> "") and _
					   (trim(Request.Form("campo5_ini_a")) <> "") then
						sCampo5 = trim(Request.Form("cb_campo5")) & "|" & _
								  trim(Request.Form("campo5_ini_d")) & "/" & _
								  trim(Request.Form("campo5_ini_m")) & "/" & _
								  trim(Request.Form("campo5_ini_a")) & "|" & _
								  trim(Request.Form("campo5_ini_d")) & "/" & _
								  trim(Request.Form("campo5_ini_m")) & "/" & _
								  trim(Request.Form("campo5_ini_a"))
					else
						sCampo5 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo5")) = "TABELA") then
				if (trim(Request.Form("cod_tabela5")) <> "") and (trim(Request.Form("cb_campo5")) <> "") and _ 
				   (trim(Request.Form("cb_campo5")) <> "0") then
					sCampo5 = trim(Request.Form("cod_tabela5")) & "," & trim(Request.Form("cb_campo5"))
				end if
			else
				sCampo5 = ""
			end if
		else
			if (trim(Request.Form("campo5")) <> "") then
				sCampo5 = trim(Request.Form("campo5"))
			else
				sCampo5 = ""
			end if
		end if

		if (trim(Request.Form("tipo_campo6")) <> "") then
			sTipoCampo6 = trim(Request.Form("tipo_campo6"))
			if (trim(Request.Form("tipo_campo6")) = "ALFA") then
				sCampo6 = trim(Request.Form("campo6"))
			elseif (trim(Request.Form("tipo_campo6")) = "LOGICO") then
				sCampo6 = trim(Request.Form("campo6"))
			elseif (trim(Request.Form("tipo_campo6")) = "NUM") then
				if (trim(Request.Form("cb_campo6")) = "3") then
					sCampo6 = trim(Request.Form("cb_campo6")) & "|" & _
					          trim(Request.Form("campo6_ini")) & "|" & _
					          trim(Request.Form("campo6_fim"))
				else
					if trim(Request.Form("campo6_ini")) <> "" then
						sCampo6 = trim(Request.Form("cb_campo6")) & "|" & _
								  trim(Request.Form("campo6_ini")) & "|" & _
								  trim(Request.Form("campo6_ini"))
					else
						sCampo6 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo6")) = "DATA") then
				if (trim(Request.Form("cb_campo6")) = "3") then
					if (trim(Request.Form("campo6_ini_d")) <> "") and (trim(Request.Form("campo6_ini_m")) <> "") and _
					   (trim(Request.Form("campo6_ini_a")) <> "") and (trim(Request.Form("campo6_fim_d")) <> "") and _
                       (trim(Request.Form("campo6_fim_m")) <> "") and (trim(Request.Form("campo6_fim_a")) <> "") then
                        sCampo6 = trim(Request.Form("cb_campo6")) & "|" & _
					          trim(Request.Form("campo6_ini_d")) & "/" & _
					          trim(Request.Form("campo6_ini_m")) & "/" & _
					          trim(Request.Form("campo6_ini_a")) & "|" & _
					          trim(Request.Form("campo6_fim_d")) & "/" & _
					          trim(Request.Form("campo6_fim_m")) & "/" & _
					          trim(Request.Form("campo6_fim_a"))
                    else
						sCampo6 = ""
					end if
				else
					if (trim(Request.Form("campo6_ini_d")) <> "") and (trim(Request.Form("campo6_ini_m")) <> "") and _
					   (trim(Request.Form("campo6_ini_a")) <> "") then
						sCampo6 = trim(Request.Form("cb_campo6")) & "|" & _
								  trim(Request.Form("campo6_ini_d")) & "/" & _
								  trim(Request.Form("campo6_ini_m")) & "/" & _
								  trim(Request.Form("campo6_ini_a")) & "|" & _
								  trim(Request.Form("campo6_ini_d")) & "/" & _
								  trim(Request.Form("campo6_ini_m")) & "/" & _
								  trim(Request.Form("campo6_ini_a"))
					else
						sCampo6 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo6")) = "TABELA") then
				if (trim(Request.Form("cod_tabela6")) <> "") and (trim(Request.Form("cb_campo6")) <> "") and _ 
				   (trim(Request.Form("cb_campo6")) <> "0") then
					sCampo6 = trim(Request.Form("cod_tabela6")) & "," & trim(Request.Form("cb_campo6"))
				end if
			else
				sCampo6 = ""
			end if
		else
			if (trim(Request.Form("campo6")) <> "") then
				sCampo6 = trim(Request.Form("campo6"))
			else
				sCampo6 = ""
			end if
		end if

		if (trim(Request.Form("tipo_campo7")) <> "") then
			sTipoCampo7 = trim(Request.Form("tipo_campo7"))
			if (trim(Request.Form("tipo_campo7")) = "ALFA") then
				sCampo7 = trim(Request.Form("campo7"))
			elseif (trim(Request.Form("tipo_campo7")) = "LOGICO") then
				sCampo7 = trim(Request.Form("campo7"))
			elseif (trim(Request.Form("tipo_campo7")) = "NUM") then
				if (trim(Request.Form("cb_campo7")) = "3") then
					sCampo7 = trim(Request.Form("cb_campo7")) & "|" & _
					          trim(Request.Form("campo7_ini")) & "|" & _
					          trim(Request.Form("campo7_fim"))
				else
					if trim(Request.Form("campo7_ini")) <> "" then
						sCampo7 = trim(Request.Form("cb_campo7")) & "|" & _
								  trim(Request.Form("campo7_ini")) & "|" & _
								  trim(Request.Form("campo7_ini"))
					else
						sCampo7 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo7")) = "DATA") then
				if (trim(Request.Form("cb_campo7")) = "3") then
					if (trim(Request.Form("campo7_ini_d")) <> "") and (trim(Request.Form("campo7_ini_m")) <> "") and _
					   (trim(Request.Form("campo7_ini_a")) <> "") and (trim(Request.Form("campo7_fim_d")) <> "") and _
                       (trim(Request.Form("campo7_fim_m")) <> "") and (trim(Request.Form("campo7_fim_a")) <> "") then
                        sCampo7 = trim(Request.Form("cb_campo7")) & "|" & _
					          trim(Request.Form("campo7_ini_d")) & "/" & _
					          trim(Request.Form("campo7_ini_m")) & "/" & _
					          trim(Request.Form("campo7_ini_a")) & "|" & _
					          trim(Request.Form("campo7_fim_d")) & "/" & _
					          trim(Request.Form("campo7_fim_m")) & "/" & _
					          trim(Request.Form("campo7_fim_a"))
                    else
						sCampo7 = ""
					end if
				else
					if (trim(Request.Form("campo7_ini_d")) <> "") and (trim(Request.Form("campo7_ini_m")) <> "") and _
					   (trim(Request.Form("campo7_ini_a")) <> "") then
						sCampo7 = trim(Request.Form("cb_campo7")) & "|" & _
								  trim(Request.Form("campo7_ini_d")) & "/" & _
								  trim(Request.Form("campo7_ini_m")) & "/" & _
								  trim(Request.Form("campo7_ini_a")) & "|" & _
								  trim(Request.Form("campo7_ini_d")) & "/" & _
								  trim(Request.Form("campo7_ini_m")) & "/" & _
								  trim(Request.Form("campo7_ini_a"))
					else
						sCampo7 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo7")) = "TABELA") then
				if (trim(Request.Form("cod_tabela7")) <> "") and (trim(Request.Form("cb_campo7")) <> "") and _ 
				   (trim(Request.Form("cb_campo7")) <> "0") then
					sCampo7 = trim(Request.Form("cod_tabela7")) & "," & trim(Request.Form("cb_campo7"))
				end if
			else
				sCampo7 = ""
			end if
		else
			if (trim(Request.Form("campo7")) <> "") then
				sCampo7 = trim(Request.Form("campo7"))
			else
				sCampo7 = ""
			end if
		end if

    	if (trim(Request.Form("tipo_campo8")) <> "") then
			sTipoCampo8 = trim(Request.Form("tipo_campo8"))
			if (trim(Request.Form("tipo_campo8")) = "ALFA") then
				sCampo8 = trim(Request.Form("campo8"))
			elseif (trim(Request.Form("tipo_campo8")) = "LOGICO") then
				sCampo8 = trim(Request.Form("campo8"))
			elseif (trim(Request.Form("tipo_campo8")) = "NUM") then
				if (trim(Request.Form("cb_campo8")) = "3") then
					sCampo8 = trim(Request.Form("cb_campo8")) & "|" & _
					          trim(Request.Form("campo8_ini")) & "|" & _
					          trim(Request.Form("campo8_fim"))
				else
					if trim(Request.Form("campo8_ini")) <> "" then
						sCampo8 = trim(Request.Form("cb_campo8")) & "|" & _
								  trim(Request.Form("campo8_ini")) & "|" & _
								  trim(Request.Form("campo8_ini"))
					else
						sCampo8 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo8")) = "DATA") then
				if (trim(Request.Form("cb_campo8")) = "3") then
					if (trim(Request.Form("campo8_ini_d")) <> "") and (trim(Request.Form("campo8_ini_m")) <> "") and _
					   (trim(Request.Form("campo8_ini_a")) <> "") and (trim(Request.Form("campo8_fim_d")) <> "") and _
                       (trim(Request.Form("campo8_fim_m")) <> "") and (trim(Request.Form("campo8_fim_a")) <> "") then
                        sCampo8 = trim(Request.Form("cb_campo8")) & "|" & _
					          trim(Request.Form("campo8_ini_d")) & "/" & _
					          trim(Request.Form("campo8_ini_m")) & "/" & _
					          trim(Request.Form("campo8_ini_a")) & "|" & _
					          trim(Request.Form("campo8_fim_d")) & "/" & _
					          trim(Request.Form("campo8_fim_m")) & "/" & _
					          trim(Request.Form("campo8_fim_a"))
                    else
						sCampo8 = ""
					end if
				else
					if (trim(Request.Form("campo8_ini_d")) <> "") and (trim(Request.Form("campo8_ini_m")) <> "") and _
					   (trim(Request.Form("campo8_ini_a")) <> "") then
						sCampo8 = trim(Request.Form("cb_campo8")) & "|" & _
								  trim(Request.Form("campo8_ini_d")) & "/" & _
								  trim(Request.Form("campo8_ini_m")) & "/" & _
								  trim(Request.Form("campo8_ini_a")) & "|" & _
								  trim(Request.Form("campo8_ini_d")) & "/" & _
								  trim(Request.Form("campo8_ini_m")) & "/" & _
								  trim(Request.Form("campo8_ini_a"))
					else
						sCampo8 = ""
					end if
				end if
			elseif (trim(Request.Form("tipo_campo8")) = "TABELA") then
				if (trim(Request.Form("cod_tabela8")) <> "") and (trim(Request.Form("cb_campo8")) <> "") and _ 
				   (trim(Request.Form("cb_campo8")) <> "0") then
					sCampo8 = trim(Request.Form("cod_tabela8")) & "," & trim(Request.Form("cb_campo8"))
				end if
			else
				sCampo8 = ""
			end if
		else
			if (trim(Request.Form("campo8")) <> "") then
				sCampo8 = trim(Request.Form("campo8"))
			else
				sCampo8 = ""
			end if
		end if

		if (trim(Request.Form("campo_ordenacao")) <> "") then
			CodigoCampoOrdenacao = CLng(Request.Form("campo_ordenacao"))
		else
			CodigoCampoOrdenacao = 0
		end if

		if (trim(Request.Form("campo_ordem")) <> "") then
			OrdemPesquisa = CInt(Request.Form("campo_ordem"))
		else
			OrdemPesquisa = 0
		end if

		sURL_VOLTAR = "?ajax=1&dados=" & Server.UrlEncode(sDados) & "&objeto=" & iObjeto & "&contexto=" & iContexto & "&imagem=" & iImagem & _
                      "&campo1=" & Server.URLEncode(sCampo1) & "&campo2=" & Server.URLEncode(sCampo2) & "&campo3=" & Server.URLEncode(sCampo3) & "&campo4=" & Server.URLEncode(sCampo4) & _
                      "&campo5=" & Server.URLEncode(sCampo5) & "&campo6=" & Server.URLEncode(sCampo6) & "&campo7=" & Server.URLEncode(sCampo7) & "&campo8=" & Server.URLEncode(sCampo8)
	end if

	if (sURL_VOLTAR <> "") then
		sURL_VOLTAR = sURL_VOLTAR & _
			"&campo_ordenacao=" & CodigoCampoOrdenacao & "&campo_ordem=" & OrdemPesquisa 
	end if

	'Response.Write "Campo 1 => " & sCampo1 & "<br>"
	'Response.Write "Campo 2 => " & sCampo2 & "<br>"
	'Response.Write "Campo 3 => " & sCampo3 & "<br>"
	'Response.Write "Campo 4 => " & sCampo4 & "<br>"

	Set objPesquisa = ROServer.CreateComplexType("TMensagem")
	Set ParamPesq = ROServer.CreateComplexType("TPesquisa")
	
	ParamPesq.sPalavraChave = sDados
	ParamPesq.iMaterial     = iObjeto
	ParamPesq.iContexto 	= iContexto
	ParamPesq.iImagem   	= iImagem
    ParamPesq.OPERADOR    	= Session("codigo") 
	ParamPesq.sCampo1		= sCampo1
	ParamPesq.sCampo2		= sCampo2
	ParamPesq.sCampo3		= sCampo3
	ParamPesq.sCampo4		= sCampo4
	ParamPesq.sCampo5		= sCampo5
	ParamPesq.sCampo6		= sCampo6
	ParamPesq.sCampo7		= sCampo7
	ParamPesq.sCampo8		= sCampo8
	ParamPesq.CampoOrdenacao= CodigoCampoOrdenacao
	ParamPesq.OrdemPesquisa	= OrdemPesquisa

    Set objPesquisa = ROService.Pesquisa(ParamPesq, false)
	
	if (objPesquisa.bResult = false) then %>
        <div id="divTituloPesquisa">
			<span style="float: left">
				<a class="link_menu" href="#" onClick=NovaPesquisa()>Home</a> > Resumo
			</span>
            <!-- #include file="../asp/botaoLogin.asp" -->
		</div>
        <br><br>
        <img src='imagens/erro.gif'>&nbsp;
	    <b>
     	<%= objPesquisa.sMsg %>
        </b>
        <br><br>
        <% if sURL_VOLTAR <> "" then %>
        	<input type='button' name='revisar' size='35' value='Revisar campos pesquisados' onClick="LinkVoltarPesquisa('<%=sURL_VOLTAR%>');">
        <% end if %>
	<% else %>
        <div id="divTituloPesquisa">
			<span style="float: left">
				<a class="link_menu" href="#" onClick=NovaPesquisa()>Home</a> > Resumo
			</span>
            <!-- #include file="../asp/botaoLogin.asp" -->
		</div>
        <br><br>
        <%
		
		Set xmlDoc = CreateObject("Microsoft.xmldom")
		xmlDoc.async = False
		xmlDoc.loadxml objPesquisa.sMsg
		Set xmlRoot = xmlDoc.documentElement
		
		if xmlRoot.nodeName  = "RESULTADO" then
			sTabTemp = xmlRoot.attributes.getNamedItem("TMP").value
			iQtde    = xmlRoot.attributes.getNamedItem("QTDE").value
			
			if (CInt(iQtde) = 0) then
				Response.Write "<b>Nenhum item foi encontrado</b>"
		        Response.Write "<br><br>"
        		if sURL_VOLTAR <> "" then 
		        	Response.Write "<input type='button' name='revisar' size='35' value='Revisar campos pesquisados' onClick=""LinkVoltarPesquisa('" & sURL_VOLTAR & "');"">"
        		end if
			else				
				if (CInt(iQtde) = 1) then
					sQtd = iQtde & " item encontrado"
				else
					sQtd = iQtde & " itens encontrados"
				end if

				Response.Write "<form action='asp/resumo.asp?content=detalhe_resultados' name='frm_pesquisa' method='POST' id='frm_pesquisa' onSubmit='return ValidaForm();'>"
				Response.Write "<table width='95%' border='0' cellspacing='0' cellpadding='0'>"			
				if(CInt(iQtde) > 100)then
					Response.Write "<tr height='30'><td>:: " & iQtde & " itens encontrados</td></tr>"
				end if
				
				sFuncExibirListagem = "exibeListagem('"&sTabTemp&"',0,0,'"& _
					Server.UrlEncode(sDados)&"',"&CStr(iObjeto)&","&CStr(iContexto)&",'"& _
					Server.UrlEncode(sCampo1)&"','"&Server.UrlEncode(sCampo2)&"','"& _
					Server.UrlEncode(sCampo3)&"','"&Server.UrlEncode(sCampo4)&"','"& _
					Server.UrlEncode(sCampo5)&"','"&Server.UrlEncode(sCampo6)&"','"& _
					Server.UrlEncode(sCampo7)&"','"&Server.UrlEncode(sCampo8)&"','"& _
					CStr(CodigoCampoOrdenacao)&"','"&CStr(OrdemPesquisa)&"');"

				Response.Write "<tr><td class='esquerda'>"
				Response.Write "<input style='width: auto;' type='button' name='listar' value='"&sQtd&"' href='#' onclick="""&sFuncExibirListagem&""">"
				Response.Write "&nbsp;&nbsp;<input class='input_busca' type='text' name='dados_outro' id='dados_outro' size='25'>"
				Response.Write "&nbsp;<input type='submit' name='submit' size='35' value='Refinar'>"
				Response.Write "</td></tr>"
				Response.Write "</table>"
							
				Response.Write "<br>"

				Response.Write "<input type='hidden' name='dados' value='"&sDados&"'>"
				Response.Write "<input type='hidden' name='objeto' value='"&CStr(iObjeto)&"'>"
				Response.Write "<input type='hidden' name='contexto' value='"&CStr(iContexto)&"'>"
				Response.Write "<input type='hidden' name='imagem' value='"&CStr(iImagem)&"'>"
				Response.Write "<input type='hidden' name='campo1' value='"&sCampo1&"'>"
				Response.Write "<input type='hidden' name='campo2' value='"&sCampo2&"'>"
				Response.Write "<input type='hidden' name='campo3' value='"&sCampo3&"'>"
				Response.Write "<input type='hidden' name='campo4' value='"&sCampo4&"'>"
				Response.Write "<input type='hidden' name='campo5' value='"&sCampo5&"'>"
				Response.Write "<input type='hidden' name='campo6' value='"&sCampo6&"'>"
				Response.Write "<input type='hidden' name='campo7' value='"&sCampo7&"'>"
				Response.Write "<input type='hidden' name='campo8' value='"&sCampo8&"'>"
				Response.Write "<input type='hidden' name='campo_ordenacao' value='"&CodigoCampoOrdenacao&"'>"
				Response.Write "<input type='hidden' name='campo_ordem' value='"&OrdemPesquisa&"'>"
					
				For Each xmlPNode In xmlRoot.childNodes
					if xmlPNode.nodeName = "CONTEXTOS" then
						sDescContexto = xmlPNode.attributes.getNamedItem("DESC").value
						
						Response.Write "<table class='grid' width='95%' cellspacing=1 cellpading=0>"
						Response.Write "<tr class='tr_grid_cabecalho' style='height: 23px'>"
						Response.Write "<td class='td_grid_cabecalho esquerda'>"
						Response.Write "&nbsp;"&sDescContexto
						Response.Write "</td>"
						Response.Write "<td class='td_grid_cabecalho' style='text-align: center; width: 120px'>"
						Response.Write "Itens encontrados"
						Response.Write "</td>"
						Response.Write "</tr>"
						
						For Each xmlContexto In xmlPNode.childNodes
							if xmlContexto.nodeName = "CONTEXTO" then
								sTmpContexto = xmlContexto.attributes.getNamedItem("DESC").value
								iTmpContexto = xmlContexto.attributes.getNamedItem("COD").value
								iQtde        = xmlContexto.attributes.getNamedItem("QTDE").value
                                iTmpObjeto = xmlContexto.attributes.getNamedItem("MATERIAL").value

								sFuncExibirListagem = "exibeListagem('"&sTabTemp&"',"&iTmpObjeto&","&iTmpContexto&",'"& _
									Server.UrlEncode(sDados)&"',"&CStr(iObjeto)&","&CStr(iContexto)&",'"& _
									Server.UrlEncode(sCampo1)&"','"&Server.UrlEncode(sCampo2)&"','"& _
									Server.UrlEncode(sCampo3)&"','"&Server.UrlEncode(sCampo4)&"','"& _
									Server.UrlEncode(sCampo5)&"','"&Server.UrlEncode(sCampo6)&"','"& _
									Server.UrlEncode(sCampo7)&"','"&Server.UrlEncode(sCampo8)&"','"& _
									CStr(CodigoCampoOrdenacao)&"','"&CStr(OrdemPesquisa)&"');"
							
								Response.Write "<tr>"
								Response.Write "<td class='td_valor'>"
								Response.Write "&nbsp;<a class='link_valor' href='#' onclick=""javascript:"&sFuncExibirListagem&""">" & sTmpContexto & "</a>"
								Response.Write "</td>"
								Response.Write "<td class='td_valor' style='text-align: center'>"
								Response.Write iQtde
								Response.Write "</td>"
								Response.Write "</tr>"
							end if
						Next
						
						Response.Write "</table>"
						Response.Write "<br>"
					end if
					
					if xmlPNode.nodeName = "MATERIAIS" then
						sDescObjeto = xmlPNode.attributes.getNamedItem("DESC").value
					
						Response.Write "<table class='grid' width='95%' cellspacing=1 cellpading=0>"
						Response.Write "<tr class='tr_grid_cabecalho' style='height: 23px'>"
						Response.Write "<td class='td_grid_cabecalho esquerda'>"
						Response.Write "&nbsp;" & sDescObjeto 
						Response.Write "</td>"
						Response.Write "<td class='td_grid_cabecalho' style='text-align: center; width: 120px'>"
						Response.Write "Itens encontrados"
						Response.Write "</td>"
						Response.Write "</tr>"
						
						For Each xmlContexto In xmlPNode.childNodes
							if xmlContexto.nodeName = "MATERIAL" then
								sTmpObjeto = xmlContexto.attributes.getNamedItem("DESC").value
								iTmpObjeto = xmlContexto.attributes.getNamedItem("COD").value
								iQtde      = xmlContexto.attributes.getNamedItem("QTDE").value

								sFuncExibirListagem = "exibeListagem('"&sTabTemp&"',"&iTmpObjeto&",0,'"& _
									Server.UrlEncode(sDados)&"',"&CStr(iObjeto)&","&CStr(iContexto)&",'"& _
									Server.UrlEncode(sCampo1)&"','"&Server.UrlEncode(sCampo2)&"','"& _
									Server.UrlEncode(sCampo3)&"','"&Server.UrlEncode(sCampo4)&"','"& _
									Server.UrlEncode(sCampo5)&"','"&Server.UrlEncode(sCampo6)&"','"& _
									Server.UrlEncode(sCampo7)&"','"&Server.UrlEncode(sCampo8)&"','"& _
									CStr(CodigoCampoOrdenacao)&"','"&CStr(OrdemPesquisa)&"');"
								
								Response.Write "<tr>"
								Response.Write "<td class='td_valor'>"
								Response.Write "&nbsp;<a class='link_valor' href='#' onclick=""javascript:"&sFuncExibirListagem&""">" & sTmpObjeto & "</a>"
								Response.Write "</td>"
								Response.Write "<td class='td_valor' style='text-align: center'>"
								Response.Write iQtde
								Response.Write "</td>"
								Response.Write "</tr>"
							end if
						Next
						
						Response.Write "</table>"
					end if
				Next

				Response.Write "</form>"
			end if
		end if
		
		Set xmlRoot = nothing
		Set xmlDoc = nothing
	end if
end if	

Set ParamPesq = nothing
Set objPesquisa = nothing
	
Set ROService = nothing
Set ROServer = nothing
%>