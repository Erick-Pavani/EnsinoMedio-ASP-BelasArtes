<html>
<head></head>
<body>
<%set Banco = server.CreateObject ("ADODB.Connection")%>
<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%NomeAdm = Request.Form("txtNomeAdm")%>
<%SenhaAdm = Request.Form("txtSenhaAdm")%>
<%SQL = " Select * from Usuario where Login = '" & NomeAdm & "' And Senha = '" & SenhaAdm & "' "%>
	<%Set Tabela_Autenticacao = banco.execute(SQL)%>
<%If tabela_Autenticacao.EOF Then%>
	<%Response.Write "<script> alert ('Senha ou Login inválidos!');window.history.go(-1); </script>"%>
<%Else%>
	<%Response.Redirect ("LayoutAdm.Asp")%>
<%End If%> 
</body>
</html>