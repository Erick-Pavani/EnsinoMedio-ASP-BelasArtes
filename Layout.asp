<html>
<head></head>
<body bgcolor="FFCC99">
<table border =1 width = 100% height = 100%>
<tr>
<td height = 10% colspan=2>
<!--#Include file ="Cabecalho.asp"--> 
</tr>
<tr>
<td width=15%><!--#Include file ="Menu.asp"--> 
<td Width=85%><%If request("Pagina") = "" Then%>
<!--#Include file ="Vitrine.asp"-->
<%ElseIf request("Pagina") = "Procura" then%>
<!--#Include file ="Procura.asp"-->
<%ElseIf request("Pagina") = "Pintores" Then%>
<!--#Include file ="Pintores.asp"-->
<%ElseIf request("Pagina") = "DetalhePintores" Then%>
<!--#Include file ="DetalhePintores.asp"-->
<%ElseIf request("Pagina") = "DetalheObra" Then%>
<!--#Include file ="DetalheObra.asp"-->
<%ElseIf request("Pagina") = "Promocoes" Then%>
<!--#Include file ="Promocoes.asp"-->
<%ElseIf request("Pagina") = "ComoComprar" Then%>
<!--#Include file ="ComoComprar.asp"-->
<%ElseIf request("Pagina") = "Contato" Then%>
<!--#Include file ="Contato.Asp"-->
<%ElseIf request("Pagina") = "Cadastro" Then%>
<!--#Include file ="Cadastro.Asp"-->
<%ElseIf request("Pagina") = "RespostaCadastro" Then%>
<!--#Include file ="RespostaCadastro.Asp"-->
<%ElseIf request("Pagina") = "RespostaProcura" Then%>
<!--#Include file ="RespostaProcura.Asp"-->
<%ElseIf request("Pagina") = "Obras" Then%>
<!--#Include file ="Obras.Asp"-->
<%ElseIf request("Pagina") = "FecharPedido" Then%>
<!--#Include file ="FecharPedido.Asp"-->
<%End If%> 
</body>
</html>