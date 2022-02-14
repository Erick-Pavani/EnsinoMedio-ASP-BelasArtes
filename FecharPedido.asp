<htmL>
<head></head>
<body>
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
		<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%If Session("Pedido").count = 0 then%>
	<%response.write ("<p align=center><font face=verdana size=3 color=red> Não há produtos em sua cesta</font>")%>
<%Else%>
<%Response.Write "<tr><td>&nbsp;</td><td>Total:</td><td colspan=3><font face=verdana size=2>" & FormatCurrency(total,2) & "</font></td></tr>"%>
<%Set Produto = nothing%>
<form action = atualizar_carrinho.asp method=post>
  <table border=1 align=center width=80%>
    <tr> 
      <th colspan=4> 
        <h3>Suas compras</h3>
      </th>
    </tr>
    <tr> 
      <td align=center>Excluir </td>
      <td align=center>Quantidade </td>
      <td align=center>Obra</td>
      <td align=center>Valor</td>
    </tr>
<%total = 0%>

<%for each Produto in Session("Pedido")%>
		<%SQL = "SELECT * FROM obra WHERE CodObra =" & Produto  & ""%>
		<%Set rstPedido = Banco.execute(SQL)%>
		<%If not  rstpedido.eof Then%>
<%Response.Write "<tr><td><input type=checkbox name=exclui" & Produto &" value = 1></td><td><input type=text size=2 name=quantidade" & Produto & " value=" & Session("Pedido").item(produto) &"></td><td><font face=verdana size=2>"& rstPedido("Obra") & "</font></td><td><font face=verdana size=2>" & FormatCurrency(rstPedido("Preco"),2)  & "</td></font></tr><br>"%>
	<%if Session("Pedido").item(Produto) > 1 Then%>
				<%total = total + Session("Pedido").item(Produto) * rstPedido("Preco")%>
			<%else%>
				<%total = total + rstPedido("Preco")%>
			<%end if%>
		<%end if%>
<%next%> 
<td colspan=4>
        <INPUT TYPE="submit" value="Atualizar">
        <INPUT TYPE="submit" name=acao value="Fechar Pedido">
</td>
  </table>
</form>
<%Banco.close%>
<%Set Banco = nothing%>
<%End If%>
</body>
</html>