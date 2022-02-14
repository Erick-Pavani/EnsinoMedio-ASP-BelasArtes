<html>
<head></head>
<body>
<%set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
	<table border = "2" height = 100% width = 100%><br></br>
	<center><font size="8" color="Red">Promocoes</font></center>
<%SQL = "Select * From Obra Where Flag_Promocao = True"%>
<%Set Tabela = Banco.Execute(SQL)%>
	<%If Tabela.EOF Then%>
		<%Response.Write "Arquivo não encontrado!"%>
	<%Else%>	
		<% Do While Not Tabela.EOF %>
		<tr>
		<td>
		<center>
			<a href = "Layout.Asp?Pagina=DetalheObra&CodObra=<%Response.Write Tabela("CodObra")%>"><img SRC = "Imagens\obras\<%Response.Write Tabela("Foto")%>_p.jpg"><br><br>
			<%Response.Write Tabela("Obra")%></a>
		</center>
		</td>
			<%Tabela.MoveNext%>
	<%If Tabela.EOF Then%>		
		<td>
		</td>
	<%Else%>
		<td>
		<center>
			<a href = "Layout.Asp?Pagina=DetalheObra&CodObra=<%Response.Write Tabela("CodObra")%>"><img SRC = "Imagens\obras\<%Response.Write Tabela("Foto")%>_p.jpg"><br><br>
			<%Response.Write Tabela("Obra")%></a>
		</td>
		</center>
			<%Tabela.MoveNext%>
		</tr>
	<%End If%>
		<%Loop%>
	<%End If%>
</body>
</html>