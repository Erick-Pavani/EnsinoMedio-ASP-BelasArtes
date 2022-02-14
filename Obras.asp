<html>
<head></head>
<body>
<%Set Banco = server.CreateObject ("ADODB.Connection")%>
	<%Banco.open "provider = Microsoft.jet.OLEDB.4.0;Data Source= "& server.mappath("belasartes.mdb")%>
<%
Set rstObras = Server.CreateObject("ADODB.RecordSet")
rstObras.Pagesize = 4
rstObras.Open "SELECT Obra.CodObra, Obra.Foto, Obra.Obra from Obra", Banco, 3 , 3

if  request.querystring("Pag")= "" then
	intpagina = 1
else
	if  Cint(request.querystring("Pag")) < 1 then
		intpagina = 1
	else
		if  Cint(request.querystring("Pag")) > rstObras.Pagecount then
			intpagina = rstObras.Pagecount
		else
			intpagina = Cint(request.querystring("Pag"))
		end if
	end if
end if

rstObras.AbsolutePage = intpagina
%>

<table border=1 align=center width="100%">
  <tr> 
    <td colspan="2">Obras a Venda </td>
  </tr>
  <% intRec= 0
While  intRec < rstObras.pagesize AND Not rstObras.eof %> 
  <tr> 
    <td valign=boton width=50%> 
 <a href="Layout.Asp?Pagina=DetalheObra&CodObra=<%=rstObras("CodObra")%>&Pag=<%=intpagina%>"><img src="imagens/obras/<%=rstObras("Foto")%>_p.jpg" border=0> <br>
        	<%=rstObras("Obra")%></a>
    </td>
    <td valign=boton width=50%> <font size="2" face="Verdana"> 
<%rstObras.Movenext
	intRec = intRec + 1
	if rstobras.eof then
		response.write "&nbsp;"
	else%> 
<a href="Layout.Asp?Pagina=DetalheObra&CodObra=<%=rstObras("CodObra")%>&Pag="<%=intpagina%>><img src="imagens/obras/<%=rstObras("Foto")%>_p.jpg" border=0> <br>
        		<%=rstObras("Obra")%></a>
<%rstObras.Movenext
		intRec = intRec + 1
	end if%>
    </td>
<%Wend%>
<%A = 1%>
</tr>
</table>
<%For A = 1 To rstObras.Pagecount%>
<a href="Layout.Asp?Pagina=Obras&Pag=<%=A%>"><%=A%>&nbsp;</a>
<%Next%>
</body>
</html>