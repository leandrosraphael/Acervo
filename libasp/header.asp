<% Server.ScriptTimeout = 180000 %>

<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->

<%
If Session("loginext") <> "" then
	codigo = Session("loginext")
	if len(trim(Session("loginext"))) > 0 then		

		Set objValidaUsuario = ROService.ValidaLoginUsuario(codigo, Senha, true)
 
		logou = objValidaUsuario.Logado
		
 		if logou then
            nome = objValidaUsuario.Nome
            Session("logado") = "true"
            Session("nome") = objValidaUsuario.Nome
            Session("codigo") = objValidaUsuario.Codigo
        else
            mensagem = objValidaUsuario.Mensagem
			Session("loginext") = ""
        end if	
	else
		Session("loginext") = ""
	end if
end if
 %>	
