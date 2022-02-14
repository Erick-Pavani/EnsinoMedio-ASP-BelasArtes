<html>
<head></head>
<body>
<%Set Banco = Server.CreateObject ("ADODB.Connection")%>
		<%Banco.Open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%Procura = Request.Form("txtProcura")%>
<%SQL = "Select Obra.CodObra,Obra.ObraDescricao,Obra.Obra,PintorNome from Obra inner join Pintor on Obra.CodPintor = Pintor.CodPintor where Obra.ObraDescricao like '%" & Procura & "%' or Obra.Obra like '%" & Procura & "%' or  Pintor.PintorNome like '%" & Procura & "%'" %>
<%Set Tabela = Banco.Execute(SQL)%>
<%Do While Not Tabela.EOF%>
<center><a href = "Layout.Asp?Pagina=DetalheObra&CodObra=<%Response.Write Tabela("CodObra")%>"><%Response.Write Tabela("Obra")%></center></a><br>
<center><%response.write Tabela("ObraDescricao")%></center><br><br>
<%Tabela.movenext%>
<%Loop%>
</body>
</html> 	