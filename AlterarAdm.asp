<html>
<head></head>
<body>
<center><font color = "Red" size = "7"> Obras a alterar: </font></center><br><br> 
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
		<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%SQL = "Select * From Obra"%>
<%Set Tabela = Banco.Execute(SQL)%>
	<%If Tabela.EOF Then%>
		<%Response.Write "Arquivo não encontrado!"%>
	<%Else%>	
		<% Do While Not Tabela.EOF %>
			<center><a href ="LayoutAdm.Asp?Pagina=DetalheAlterar&CodObra=<%Response.Write Tabela("CodObra")%>&CodPintor=<%Response.Write Tabela("CodPintor")%>"><%Response.Write Tabela("Obra")%></a><br></center>
			<%Tabela.MoveNext%>
		<%Loop%>
	<%End If%>
</body>
</html>