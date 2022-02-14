<% Codigo = CINT (Request.QueryString("CodObra"))%>
<%If Session("Pedido").Exists(Codigo) Then%>
	<%Response.Write "<script> alert ('Este quadro j? se encontra no seu carrinho!');window.history.go(-2); </script>"%>
<%Else%>
	<%Session("Pedido").Add Codigo,1%>
	<%Response.Redirect "Layout.Asp"%>
<%End If%>