<% Server.ScriptTimeout = 180000 %>

<!-- #include file="../config.asp" -->
<!-- #include file="../libasp/roclient.asp" -->

<%
    loginUsuario = Request("login")
	senha = Request("senha")
       
    if len(trim(loginUsuario)) > 0 then

		Set objValidaUsuario = ROService.ValidaLoginUsuario(loginUsuario, Senha, false)
 
        logou = objValidaUsuario.Logado
		
 		if logou then
            nome = objValidaUsuario.Nome
            Session("logado") = "true"
            Session("nome") = objValidaUsuario.Nome
            Session("codigo") = objValidaUsuario.Codigo
            
            Response.ContentType = "application/json"
            Response.Write("{ ""logou"":""true"",""nome"":"""&nome&"""}")
        else
            mensagem = objValidaUsuario.Mensagem

            Response.ContentType = "application/json"
            Response.Write("{ ""logou"":""false"",""mensagem"":"""&mensagem&"""}")
        end if
        
        Response.Flush

    end if
%>