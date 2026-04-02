<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<!-- #include file="../libasp/funcoes.asp" -->
<html>
	<head>
		<meta http-equiv="content-type" content="text/html;charset=UTF-8" />
		<meta http-equiv="X-UA-Compatible" content="IE=edge" />
		<title></title>
	</head>
<%
	iIndexSrv = IntQueryString("servidor", 1)
	global_idioma = IntQueryString("iIdioma", 0)
%>
	<frameset rows="60,*" framespacing="0" frameborder="0">
		<frame name="imp_topFrame" src="impRec_topo.asp?iIdioma=<%=global_idioma%>&iBanner=<%=global_tipo_banner%>">
		<frame name="imp_impFrame" src="impRec.asp?servidor=<%=iIndexSrv%>&iBanner=<%=global_tipo_banner%>&iIdioma=<%=global_idioma%>">
	<noframes></noframes></frameset>
</html>