<html>
<head></head>
<body>
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
		<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<center><font color = "Red" size = "7">Obras a Excluir:</font></center><br><br> 
<%SQL = "Select * From Obra"%>
<%Set Tabela = Banco.Execute(SQL)%>
	<%If Tabela.EOF Then%>
		<%Response.Write "Arquivo não encontrado!"%>
	<%Else%>	
		<% Do While Not Tabela.EOF %>
			<center><a href = "LayoutAdm.Asp?Pagina=DetalheExcluir&CodObra=<%Response.Write Tabela("CodObra")%>&CodPintor=<%Response.Write Tabela("CodPintor")%>"><img SRC = "Imagens\obras\<%Response.Write Tabela("Foto")%>_p.jpg">
			<center><a href = "LayoutAdm.Asp?Pagina=DetalheExcluir&CodObra=<%Response.Write Tabela("CodObra")%>&CodPintor=<%Response.Write Tabela("CodPintor")%>"><%Response.Write Tabela("Obra")%></a><br><br></center>
			<%Tabela.MoveNext%>
		<%Loop%>
	<%End If%>
</body>
</html>