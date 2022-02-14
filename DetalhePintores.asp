<html>
<head></head>
<body>
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%SQL = "Select PintorNome From Pintor Where CodPintor = " & request.QueryString("CodPintor") & ""  %>
<%Set Tabela = Banco.Execute(SQL)%> 
<center>Obras De: <% Response.Write Tabela("PintorNome")%><br></center>
<% SQL = "Select * From Obra Where CodPintor = " & request.QueryString("CodPintor") & ""%>
<%Set Tabela = Banco.Execute(SQL)%>
	<%Do While Not Tabela.EOF%>
		<center><a href = "Layout.Asp?Pagina=DetalheObra&CodObra=<%Response.Write Tabela("CodObra")%>"><%Response.Write Tabela("Obra")%></a></br></center>
		<%Tabela.Movenext%>
	<%Loop%>
</body>
</html>