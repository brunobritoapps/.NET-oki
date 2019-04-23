<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Sub getBonus()
		Dim sql, arr, intarr, i
		
		sql = "SELECT [cod_bonus] " & _
					  ",[descricao] " & _
					  ",[validade] " & _
					  ",[moeda] " & _
					  ",[aplicacao] " & _
					  ",[data_inicio_contabilizacao] " & _
				  "FROM [marketingoki2].[dbo].[Cadastro_Bonus] where aplicacao = 'PONTO'"
				  
		Call search(sql ,arr, intarr)				
		If intTransp > -1 Then
			Response.Write "<input type=""hidden"" name=""hiddenIntBonus"" value="&intarr + 1&" />"
			For i=0 To intarr
				If i Mod 2 = 0 Then
					Response.Write "<tr>"
					Response.Write "<td class=""classColorRelPar""><input type=""radio"" name=""bonus"" value="""&arr(0,i)&""" OnClick=""updateBonus()"" /></td>"
					Response.Write "<td class=""classColorRelPar"" width=""15%"">"&arr(0,i)&"</td>"
					Response.Write "<td class=""classColorRelPar"" width=""4%"">"&arr(2,i)&"</td>"
					Response.Write "<td class=""classColorRelPar"" width=""4%"">"&arr(3,i)&"</td>"
					Response.Write "<td class=""classColorRelPar"">"&arr(1,i)&"</td>"
					Response.Write "</tr>"
				Else
					Response.Write "<tr>"
					Response.Write "<td class=""classColorRelImpar""><input type=""radio"" name=""bonus"" value="""&arr(0,i)&""" OnClick=""updateBonus()"" /></td>"
					Response.Write "<td class=""classColorRelImpar"">"&arr(0,i)&"</td>"
					Response.Write "<td class=""classColorRelImpar"">"&arr(2,i)&"</td>"
					Response.Write "<td class=""classColorRelImpar"">"&arr(3,i)&"</td>"
					Response.Write "<td class=""classColorRelImpar"">"&arr(1,i)&"</td>"
					Response.Write "</tr>"
				End If	
			Next
		End If
	End Sub
%>
<html>
<head>
<script>
	function updateBonus() {
		var error = 0;
		var valor = null;
		for (var i=0; i < parseInt(document.frmsearchbonuscliente.hiddenIntBonus.value); i++) {
			if (parseInt(document.frmsearchbonuscliente.hiddenIntBonus.value) == 1) {
				if (!document.frmsearchbonuscliente.bonus.checked) {
					error++;
				} else {
					valor = document.frmsearchbonuscliente.bonus.value;	
				}
			} else {
				if (!document.frmsearchbonuscliente.bonus[i].checked) {
					error++;
				}	else {
					valor = document.frmsearchbonuscliente.bonus[i].value;	
				}
			}
		}
		if (error == parseInt(document.frmsearchbonuscliente.hiddenIntBonus.value)) {
			alert("Por favor escolha um Bônus");
			return;
		} else {
			window.opener.frmPontoColetaAdm.cbBonus.value = valor;		
			window.close();
		}
	}
	
</script>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/geral.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="container">
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td id="conteudo">
					<div style="overflow:scroll;width:600px;height:317px;">
						<form action="frmsearchbonuscliente.asp" name="frmsearchbonuscliente" method="POST">
						<table cellspacing="1" cellpadding="1" width="600" id="tablelisttransportadoras">
							<tr>
								<td id="explaintitle" colspan="5" align="center">Busca de Bônus</td>
							</tr>
							<tr>
								<th width="5%" ><img src="img/check.gif"></th>
								<th>Cód. Bônus</th>
								<th>Validade</th>
								<th>Moeda</th>
								<th>Descrição</th>
							</tr>
							<%Call getBonus()%>
						</table>
						</form>
					</div>
				</td>
			</tr>
		</table>
	</div>
</div>
</body>
</html>
<%Call close()%>
