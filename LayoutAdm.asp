<html>
<head></head>
<body bgcolor="FFCC99">
<table border =2 width = 100% height = 100%>
<tr>
<td height = 10% colspan=2>
<!--#Include file ="CabecalhoAdm.asp"--> 
</tr>
<tr>
<td width=15%><!--#Include file ="MenuAdm.asp"--> 
<td Width=85%><%If request("Pagina") = "" Or request("Pagina") = "Alterar" then%>
<!--#Include file ="AlterarAdm.asp"-->
<%ElseIf request("Pagina") = "Incluir" then%>
<!--#Include file ="IncluirAdm.asp"-->
<%ElseIf request("Pagina") = "Excluir" then%>
<!--#Include file ="ExcluirAdm.asp"-->
<%ElseIf request("Pagina") = "EnviarEmail" then%>
<!--#Include file ="EnviarEmailAdm.asp"-->
<%ElseIf request("Pagina") = "RespostaIncluirAdm" then%>
<!--#Include file ="RespostaIncluirAdm.asp"-->
<%ElseIf request("Pagina") = "RespostaEnviarEmailAdm" then%>
<!--#Include file ="RespostaEnviarEmailAdm.asp"-->
<%ElseIf request("Pagina") = "DetalheExcluir" then%>
<!--#Include file ="DetalheExcluirAdm.asp"-->
<%ElseIf request("Pagina") = "RespostaExcluir" then%>
<!--#Include file ="RespostaExcluirAdm.asp"-->
<%ElseIf request("Pagina") = "DetalheAlterar" Then%>
<!--#Include file ="DetalheAlterarAdm.asp"-->
<%ElseIf request("Pagina") = "RespostaAlterar" Then%>
<!--#Include file ="RespostaAlterarAdm.asp"-->
<%End If%>
</table>
</body>
</html>