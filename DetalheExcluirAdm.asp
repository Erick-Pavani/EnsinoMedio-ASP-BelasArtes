<html>
<head></head>
<body>
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%SQL = "Select * From Obra Where CodObra ="& Request.QueryString ("CodObra")%>
<%Set Tabela = Banco.Execute(SQL)%>
<%SQL1 = "Select PintorNome From Pintor Where CodPintor=" & Request.QueryString ("CodPintor")%>
<%Set Tabela1 = Banco.Execute(SQL1)%>
<Form Action = "LayoutAdm.Asp?Pagina=RespostaExcluir&CodObra=<%Response.Write Tabela("CodObra")%>" Method = "Post">
<table border = "2" width = 100% height = 100%><br>
<center><font color = "Red" size = "6">Detalhes da Obra</font></center><br>
<tr>
<td><center><font size = "5">Nome da obra:</center></td>
<td><center><font size = "5"><%Response.Write Tabela("Obra")%></font></center></td>
</tr>
<tr>
<td><center><font size = "5">Descrição da obra:</center></td>
<td><center><font size = "5"><%Response.Write Tabela("ObraDescricao")%></font></center></td>
</tr>
<tr>
<td><center><font size = "5">Preço da obra:</center></td>
<td><center><font size = "5"><%Response.Write formatcurrency(Tabela("Preco"))%></font></center></td>
</tr>
<tr>
<td><center><font size = "5">Nome do Pintor:</center></td>
<td><center><font size = "5"><%Response.Write Tabela1("PintorNome")%></font></center></td>
</tr>
<tr>
<td><center><font size = "5">Foto:</font></center></td>
<td><center><font size = "5"><%Response.Write Tabela("Foto")%></font></center></td>
</tr>
<tr>
<td colspan = "2"><center><input type = "Submit" Value = "Excluir"></center></td>
</tr>
</body>
</html>