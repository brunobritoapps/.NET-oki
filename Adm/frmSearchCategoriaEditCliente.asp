<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Sub GetCategoriaCliente()
		Dim sSql, arrCat, intCat, i
		sSql = "SELECT * FROM Categorias"
		Call search(sSql, arrCat, intCat)
		If intCat > -1 Then
			Response.Write "<input type=""hidden"" name=""hiddenIntCat"" value="&intCat + 1&" />"
			For i=0 To intCat
				If i Mod 2 = 0 Then
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelPar'><input type=""radio"" name=""radioIDCategorias"" value="&arrCat(0,i)&" onClick=""updateCat()"" /></td>"
					Response.Write "<td class='classColorRelPar'>"&arrCat(1,i)&"</td>"
					If arrCat(2,i) = 1 Then
						Response.Write "<td class='classColorRelPar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelPar'>Não</td>"
					End If
					If arrCat(3,i) = 1 Then
						Response.Write "<td class='classColorRelPar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelPar'>Não</td>"
					End If
					Response.Write "<td class='classColorRelPar'>"&arrCat(4,i)&"</td>"
					Response.Write "</tr>"
				Else	
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelImpar'><input type=""radio"" name=""radioIDCategorias"" value="&arrCat(0,i)&" onClick=""updateCat()"" /></td>"
					Response.Write "<td class='classColorRelImpar'>"&arrCat(1,i)&"</td>"
					If arrCat(2,i) = 1 Then
						Response.Write "<td class='classColorRelImpar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelImpar'>Não</td>"
					End If
					If arrCat(3,i) = 1 Then
						Response.Write "<td class='classColorRelImpar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelImpar'>Não</td>"
					End If
					Response.Write "<td class='classColorRelImpar'>"&arrCat(4,i)&"</td>"
					Response.Write "</tr>"
				End If	
			Next
		End If
	End Sub
	
%>
<html>
<head>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/geral.css" rel="stylesheet" type="text/css">
<script>
	function updateCat() {
		 var form = document.frmSearchCategoriaEditCliente;
		 var cont = 0;
		 for (var i = 0; i < form.hiddenIntCat.value; i++) {
		 		if (form.hiddenIntCat.value == 1) {
					if (form.radioIDCategorias.checked) {window.opener.frmCadastroClienteAdm.cbCategorias.value = form.radioIDCategorias.value;}		
				} else {
					if (form.radioIDCategorias[i].checked) {window.opener.frmCadastroClienteAdm.cbCategorias.value = form.radioIDCategorias[i].value;}	
				}
		 }
		 window.close();
	}
</script>
</head>

<body>
<div id="container">
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td id="conteudo" align="left">
					<form action="frmSearchCategoriaEditCliente.asp" name="frmSearchCategoriaEditCliente" method="POST">
					<table cellpadding="1" cellspacing="1" width="500">
						<tr>
							<td id="explaintitle" align="center">Busca de Grupo de Produtos</td>
						</tr>
						<tr>
							<td>
								<div style="overflow:scroll;height:300px;width:496px;">
								<table cellpadding="1" cellspacing="1" width="480" id="tableRelCategories">
									<tr>
										<th width="5%"><img src="img/check.gif" /></th>
										<th>Descrição</th>
										<th width="10%">Ativo</th>
										<th width="20%">Coleta Domiciliar</th>
										<th width="10%">Mín. Cartuchos</th>
									</tr>
									<%Call GetCategoriaCliente()%>
								</table>
								</div>
							</td>
						</tr>
					</table>
					</form>
				</td>
			</tr>
		</table>
	</div>
</div>
</body>
</html>
<%Call close()%>
