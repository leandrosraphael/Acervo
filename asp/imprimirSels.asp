<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<!-- #include file="../libasp/funcoes.asp" -->
<html>
	<head>
		<meta http-equiv="content-type" content="text/html;charset=UTF-8" />
		<meta http-equiv="X-UA-Compatible" content="IE=edge" />
		<title></title>
	</head>
<%
	iIndexSrv = IntQueryString("Servidor", 1)
	global_idioma = IntQueryString("iIdioma", 0)
	codigoFavorito = IntQueryString("codigoFavorito", 0) 
%>
	<frameset rows="60,*" framespacing="0" frameborder="0">
		<% if Request.QueryString("Tipo") = "RB" then %>
			<frame name="imp_topFrame" src="impTop.asp?Servidor=<%=iIndexSrv%>&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>&codigoFavorito=<%=codigoFavorito%>&Tipo=RB">
			<frame name="imp_impFrame" src="refbib.asp?Servidor=<%=iIndexSrv%>&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>&codigoFavorito=<%=codigoFavorito%>">
		<% else %>
			<frame name="imp_topFrame" src="impTop.asp?Servidor=<%=iIndexSrv%>&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>&codigoFavorito=<%=codigoFavorito%>">
			<frame name="imp_impFrame" src="impSels.asp?Servidor=<%=iIndexSrv%>&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>&codigoFavorito=<%=codigoFavorito%>">
		<% end if %>
	<noframes></noframes></frameset>
</html>
