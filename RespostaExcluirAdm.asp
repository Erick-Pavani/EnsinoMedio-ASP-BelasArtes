<html>
<head></head>
<body>
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%SQL = "Delete From Obra Where CodObra =" & Request.QueryString ("CodObra")%>
<%Set Tabela = Banco.Execute(SQL)%>
<center><font color = "Red" size = "5">Obra excluída com sucesso!</font></center>
</body>
</html>