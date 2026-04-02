<%tipoIdentificador = xmlIdentificadores.attributes.getNamedItem("Tipo").value   
titleImg = xmlIdentificadores.attributes.getNamedItem("Valor").value
titleImg =  replace(replace(replace(replace(titleImg,"#URL#",""), "#HREF#", ""), "#/HREF#", ""), "#/URL#", "")
titleImg =  replace(titleImg,"#IMG#","")
urlIdentificador = xmlIdentificadores.attributes.getNamedItem("Valor").value
urlIdentificador = replace(replace(replace(replace(urlIdentificador,"#URL#","<a target='_blank' "), "#HREF#", "href='"), "#/HREF#", "'>"), "#/URL#", "</a>")
urlIdentificador = replace(urlIdentificador,"#IMG#","<img class='autImg' data-tipo='"&tipoIdentificador&"' title='"&titleImg&"'>")
identificadoresInfo = identificadoresInfo &"<td class='autImg' data-tipo='"&tipoIdentificador&"' style='display:none;'>"&urlIdentificador&"</td>"
%>