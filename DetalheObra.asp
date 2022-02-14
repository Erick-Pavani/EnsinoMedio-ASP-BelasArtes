<html>
<head></head>
<body>
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%SQL = "Select * From Obra Where CodObra =" & Request.QueryString ("CodObra")%>
<%Set Tabela = Banco.Execute(SQL)%>
<table border = "0" width = 100% height = 100%>
<tr>
<td><center>
	<img SRC = "Imagens\obras\<%Response.Write Tabela("Foto")%>_g.jpg"></center></td>
	<td><center><font color="Red" size="7">Detalhes da Obra</center><br><br><center><font color="Black" size="7"><%Response.Write Tabela("Obra")%></font></center><br>	
	<center><font color="Black" size="4"><%Response.Write Tabela("ObraDescricao")%></center><br><br><center><font color="Blue" size="5"><%Response.Write formatcurrency (tabela("Preco"))%></center><br><br><br>
	<center><a href = "Carrinho.Asp?CodObra=<%Response.Write Tabela("CodObra")%>">Comprar</a></center>
</body>
</html>