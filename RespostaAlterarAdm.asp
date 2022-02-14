<html>
<head></head>
<body>
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%Obra =  Request.Form("txtObra")%>
<%Descricao = Request.Form("txtDescricao")%>
<%Pintor = Request.Form("lstPintores")%>
<%Preco = Request.Form("txtPreco")%>
<%Vitrine = Request.Form("chkVitrine")%>
<%Promocao = Request.Form("chkPromocao")%>
<%Foto = Request.Form("txtFoto")%>
<%If Promocao = "on" Then%>
<%Promocao = "True"%>
<%Else%>
<%Promocao = "False"%>
<%End If%>
<%If Vitrine = "on" Then%>
<%Vitrine = "True"%>
<%Else%>
<%Vitrine = "False"%>
<%End If%>
<%SQL = "Update Obra Set Obra = '" & Obra & "', ObraDescricao = '" & Descricao & "', CodPintor = '" & Pintor & "', Preco = '" & Preco & "', Flag_Vitrine = " & Vitrine & ", Flag_Promocao = " & Promocao & ", Foto = '" & Foto & "' Where CodObra =" & Request.QueryString ("CodObra")%>
<%Set Tabela = Banco.Execute(SQL)%>
<center><font color = "Red" size = "5">Obra alterada com sucesso!</font><center>
</body>
</html>