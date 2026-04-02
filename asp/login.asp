<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->

<div class="popup_cabecalho_background popup_cabecalho_fonte">
    <span class="espacamento_login">Login</span>
</div>

<div class="popup">
<form name="login">
	<table class="centro" style="width: 290px; padding: 0; margin: auto;">
	<tr>
	<td style="height: 40px" class="center">
	<p class="centro">
	<%

        LoginPorMatricula = ROService.LoginPorMatricula
        
        numero_serie = ROService.GetNumeroDTA
        'Itau Paralegal
        if (numero_serie = 6209) then
           LabelLogin = "Funcional"
        else
           LabelLogin = "Matrícula"
        end if

        if (LoginPorMatricula) then
           Response.Write "Por favor, informe sua "&LabelLogin&" e senha para ter acesso aos serviços do Terminal" 
        else
    	   Response.Write "Por favor, informe seu código e senha para ter acesso aos serviços do Terminal"
        end if
        
	%>
		</p>
		</td></tr>
		<tr>
		<td style="height: 100px" class="td_center_top">	
		    <table style="width: 225px; display: inline-block;">
		        <tr>
		            <td colspan="2"><br /></td></tr><tr>
                       <input type="hidden" name="LabelLogin" value="<%=LabelLogin%>" />
                        <script>
                        <% if (LoginPorMatricula) then %>
                            var loginPorMatricula = true;
                        <% else %>
                            var loginPorMatricula = false;
                        <% end if %>
                        </script>

                        <% if (LoginPorMatricula) then 
		                      Response.Write "<td style=""width: 50px"" class=""esquerda"">"&LabelLogin&"&nbsp;</td>" %>
		                      <td style="width: 175px; height: 27px;" class="esquerda">
                                  <input type="text" class="input_login" name="loginUsuario" maxlength=252/>
                              </td>
                        <% else %> 
                              <td style="width: 50px" class="esquerda">Código&nbsp;</td> 
		                      <td style="width: 175px; height: 27px;" class="esquerda">
                                  <input type="text" class="input_login" name="loginUsuario" maxlength=9  onKeyPress="return BloqueiaNaoNumerico(event)"/>
                              </td>
                        <% end if %>

		        </tr>
		        <tr style="height: 27px;">
		            <td class="esquerda">Senha&nbsp;</td>
		            <td class="esquerda">
		                <input type="password" name="senha" class="input_login"  maxlength=10 />
		            </td>
		        </tr>
		    </table>	
		</td></tr>
		<tr>
		<td class="centro" style="height: 37px; display: inline-table;">
		    <input class="button_login" type="button" onClick="return EfetuarLogin(0);" value="login" id="button1" name="button1" />
		    &nbsp;&nbsp;<input class="button_login" type="button" onClick="FecharTelaLogin()" value="cancelar" id="button2" name="button2" />
		    <br /><br />
		</td>
		</tr>
		</table>
</form>
</div>