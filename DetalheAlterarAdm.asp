<html>
<head></head>
<body>
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%SQL = "Select * From Obra Where CodObra =" & Request.QueryString("CodObra")%>
<%Set Tabela = Banco.Execute(SQL)%>
<%SQL1 = "Select * From Pintor"%>
<%Set Tabela1 = Banco.Execute(SQL1)%>
<%CodPintor = Request.QueryString("CodPintor")%>
<Form action = "LayoutAdm.Asp?Pagina=RespostaAlterar&CodObra=<%Response.Write Tabela("CodObra")%>" Method = "Post">
<center><font color = "Red" size = "7">Atualizar Obra</font></center><br>
<table border = "2" height = 100% width = 100%>
<tr>
<td><center><font size = "4">Nome da obra:</font></center></td>
<td><center><input type = "text" size = "30" name = "txtObra" value ="<%Response.Write Tabela("Obra")%>"></center></td>
</tr> 
<tr>
<td><center><font size = "4">Descrição:</font></center></td>
<td><center><input type = "text" size = "30" name = "txtDescricao" value ="<%Response.Write Tabela("ObraDescricao")%>"></center></td>
</tr>
<tr>
<td><center><font size = "4">Pintor:</font></center></td>
<td><%If Tabela1.EOF Then%>
<%Response.write "Não existe registro!"%>
<%Else%>
<center><select name = "lstPintores">			
<%Do While Not Tabela1.EOF%>
<option value="<%Response.Write Tabela1("CodPintor")%>"<%If Tabela("CodPintor") = Tabela1("CodPintor") Then%>selected<%End If%>>
				 <%Response.Write Tabela1("PintorNome")%></option></center>
<%Tabela1.MoveNext%>
<%Loop%>
<%End If%>
</select></td>
</tr>
<tr>
<td><center><font size = "4">Preço:</font></center></td>
<td><center><input type = "text" size = "30" name = "txtPreco" value ="<%Response.Write formatcurrency (Tabela("Preco"))%>"></center></td>
</tr>
<tr>
<td><center><font size = "4">Promoção?</font></center></td>
<%If Tabela("Flag_Promocao") = True Then%>
<td><center><input type = "checkbox" name = "chkPromocao" checked></center></td>
<%Else%>
<td><center><input type = "checkbox" name = "chkPromocao"></center></td>
<%End If%>
</tr>
<tr>
<td><center><font size = "4">Vitrine?</font></center></td>
<%If Tabela("Flag_Vitrine") = True Then%>
<td><center><input type = "checkbox" name = "chkVitrine" checked></center></td>
<%Else%>
<td><center><input type = "checkbox" name = "chkVitrine"></center></td>
<%End If%>
</tr>
<tr>
<td><center><font size = "4">Foto:</font></center></td>
<td><center><input type = "text" name = "txtFoto" value = "<%Response.Write Tabela("Foto")%>"></center></td>
</tr>
<tr>
<td><center><input type = "Submit" value = "Alterar"></center></td>
<td><center><input type = "Reset" value = "Cancelar"></center></td>
</tr>
</table>
</body>
</html>