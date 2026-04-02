<%
	Class BibliotecaClass	
		'Propriedades
		Public Codigo
		Public Nome
	End Class
	
	Class ServidorClass	
		'Propriedades
		Public Nome
		Public IP
		Public Porta

		Public BibList
		
		'Construtor
		Public Sub Class_Initialize
			Set BibList = CreateObject("Scripting.Dictionary")
		End Sub
		
		'Destrutor
		Public Sub Class_Terminate
			Set BibList = nothing
		End Sub
		
		'Adicionar biblioteca
		Public Sub AddBiblioteca(Codigo, Nome)
			Set Bib = new BibliotecaClass
			
			Bib.Codigo = Codigo
			Bib.Nome = Nome
			
			BibList.Add BibList.Count, Bib
		End Sub
	End Class

	Class ServidoresClass
		Public Nome
		Public ServList
		
		'Construtor
		Public Sub Class_Initialize
			Set ServList = CreateObject("Scripting.Dictionary")
		End Sub
		
		'Destrutor
		Public Sub Class_Terminate
			Set ServList = nothing
		End Sub
		
		Public Sub AddServidor(Servidor)
			ServList.Add ServList.Count, Servidor
		End Sub
	End Class
%>