<%@ EnableSessionState=False %>

<%
    Set fileSystemObject = Server.CreateObject("Scripting.FileSystemObject")
    Set folderObject = FileSystemObject.GetFolder(Server.MapPath("../temp"))

    for each objFile in folderObject.Files
        'Apaga apenas arquivos de imagem com extensão "jpg" e que comecem com o caractere "C" (padrão da capa = "C1234.jpg")
        if ((UCase(fileSystemObject.GetExtensionName(objFile.name)) = "JPG") and (left(objFile.name, 1) = "C")) then            
            fileSystemObject.DeleteFile(Server.MapPath("../temp/" & objFile.name))    
        end if
    next
    
    Set folderObject = Nothing
    Set fileSystemObject = Nothing        
%>