<html>
<head></head>
<body>
<Form Action = "LayoutAdm.Asp?Pagina=RespostaIncluirAdm" Method = "Post">
<table border = "2" height = 100% width = 100%>
<center><font color = "Red" size = "7">Incluir Obra</font></center>
<tr>
<td><center>Nome da Obra:</center></td>
<td><input type = "text" name = "txtObra" size = "50"></td><br>
</tr><p>
<tr>
<td><center>Descrição:</center></td>
<td><input type = "text" name = "txtDescricao" size = "50"></td><br>
</tr>
<tr>
<td><center>Pintor:</center></td>
<td><%Set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%SQL = "Select * from Pintor"%>
<%Set Tabela = Banco.Execute(SQL)%> 
<%If Tabela.EOF Then%>
<%Response.write "Não existe registro!"%>
<%Else%>
<select name = "lstPintores">			
<%Do While Not tabela.EOF%>
<option value = <%response.write Tabela("CodPintor")%>><%response.write Tabela("PintorNome")%></option>
<%Tabela.Movenext%>
<%Loop%>
<%End If%>
</select>
</td>
</tr>
<tr>
<td><center>Preço: R$</center></td>
<td><input type = "text" name = "txtPreco"></td><br>
</tr>
<tr>
<td><center>Promocão?</center></td>
<td><input type = "checkbox" name = "chkPromocao"></td><br>
</tr>
<tr>
<td><center>Vitrine?</center></td>
<td><input type = "checkbox" name = "chkVitrine"></td><br>
</tr>
<tr>
<td><center>Foto:</center></td>
<td><input type = "text" name = "txtFoto"></td><br>
</tr>
</table>
<center><input type = "Submit" Value = "Incluir"> <input type = "Reset" Value = "Limpar Campos">
</body>
</html>