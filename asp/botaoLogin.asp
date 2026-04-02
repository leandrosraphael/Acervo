<% if Session("UsaLogin") = "" then %>
    <!-- #include file="../config.asp" -->
    <!-- #include file="../libasp/roclient.asp" -->    
<%
    Session("UsaLogin") = ROService.UsaLogin
 end if 

if Session("UsaLogin") = "2" then
    if Session("logado") = "true" then
        nome = Session("nome") & ","
        %> <a id="a-login" class="link_login" href="#" onClick="EfetuarLogout()">Sair</a> <%
    else
        numero_serie = ROService.GetNumeroDTA
		if (numero_serie = 6209) then
			Response.Clear
			Response.Write "<img src='imagens/erro.gif'>&nbsp;<b>Sessão expirada.</b>"
			response.End
		else
			%> <a id="a-login" class="link_login" href="#" onClick="AbrirTelaLogin()">Login</a> <%
		end if
    end if
        %> <span id="a-nome-login" class="nome_login"><%=nome%></span><%
end if
%>