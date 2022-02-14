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
<%SQL = "Insert Into Obra(Obra,ObraDescricao,CodPintor,Preco,Flag_Promocao,Flag_Vitrine,Foto) Values ('" & Obra & "','" & Descricao & "','" & Pintor & "','" & Preco & "', " & Promocao & "," & Vitrine & ", '" & Foto & "')"%>
<%Set Tabela = Banco.Execute(SQL)%>
<center><font color = "Red" size = "5">Obra inserida com sucesso!</font></center>
</body>
</html>
