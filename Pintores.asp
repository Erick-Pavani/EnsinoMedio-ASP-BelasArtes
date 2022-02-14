<html>
<body>
<center><font color="Red" size="10">Pintores</font></center><br>
<%set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%SQL = "Select * From Pintor"%>
<%set Tabela = Banco.Execute(SQL)%>
	<%If Tabela.EOF Then%>
		<%Response.Write "Registro não encontrado!"%>
	<%Else%>
		<%Do While Not Tabela.EOF%> 
			<center>
			<a href = "Layout.Asp?Pagina=DetalhePintores&CodPintor=<%Response.Write Tabela("CodPintor")%>"><%Response.Write Tabela("PintorNome")%></a><br>
			</center>
			<%Tabela.MoveNext%>
		<%Loop%>
	<%End If%>
</body>
</html>