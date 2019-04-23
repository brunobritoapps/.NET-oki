<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Sub GetGrupoProduto()
		Dim sSql, arrGroup, intGroup, i
		sSql = "SELECT * FROM Grupo_Produtos"
		Call search(sSql, arrGroup, intGroup)
		If intGroup > -1 Then
			Response.Write "<input type=""hidden"" name=""hiddenIntGrupos"" value="&intGroup + 1&" />"
			For i=0 To intGroup
				If i Mod 2 = 0 Then 
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelPar'><input type=""radio"" name=""radioIntGrupo"" value="&arrGroup(0,i)&" onClick=""updateGrupo()"" /></td>"
					Response.Write "<td class='classColorRelPar'>"&arrGroup(1,i)&"</td>"
					Response.Write "</tr>"
				Else
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelImpar'><input type=""radio"" name=""radioIntGrupo"" value="&arrGroup(0,i)&" onClick=""updateGrupo()"" /></td>"
					Response.Write "<td class='classColorRelImpar'>"&arrGroup(1,i)&"</td>"
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
	function updateGrupo() {
		 var form = document.frmSearchGroupProduto;
		 var cont = 0;
		 for (var i = 0; i < form.hiddenIntGrupos.value; i++) {
		 		if (form.hiddenIntGrupos.value == 1) {
					if (form.radioIntGrupo.checked) {window.opener.frmCadProdutosAdm.cbGrupos.value = form.radioIntGrupo.value;}		
				} else {
					if (form.radioIntGrupo[i].checked) {window.opener.frmCadProdutosAdm.cbGrupos.value = form.radioIntGrupo[i].value;}	
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
					<form action="frmSearchGroupProduto.asp" name="frmSearchGroupProduto" method="POST">
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
									</tr>
									<%Call GetGrupoProduto()%>
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
