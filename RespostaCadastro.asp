<html>
<head></head>
<body>
<%set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%Nome = Request.Form("txtNome")%>
<%Email = Request.Form("txtEmail")%>
<%SQL = "Insert Into Cadastro (Nome, Email) Values ('" & Nome & "', '" & Email & "')"%>
<% Set Tabela = Banco.Execute(SQL) %>
<center><font color = "Red" size = "5"> Você foi Cadastrado com sucesso!</font></center><br><br>
<center><a href = "Layout.Asp?Pagina="><input type= "Submit" Value= "Voltar para o site"></a></center>
</body>
</html>