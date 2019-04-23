<%
'|--------------------------------------------------------------------
'| Arquivo: MakeCookie.asp																									 
'| Autor: Jadilson
'| Data Criaчуo: 10/01/2008																					 
'| Data Modificaчуo : 10/01/2008																		 
'| Descriчуo: Cookie para resgate de bonus
'|--------------------------------------------------------------------

Dim IdProd, QtdProd

IdProd		= Request("IdProd")
QtdProd		= Request("QtdProd")

if instr(request.cookies("Okidata"), IdProd) = 0 then
	if len(trim(QtdProd)) > 0 then
		response.cookies("Okidata") = IdProd & "," & QtdProd & "@" & request.cookies("Okidata")
	end if
else
	response.cookies("Okidata") = IdProd & "," & QtdProd & "@" & request.cookies("Okidata")
	'call UpdCookie(IdProd, QtdProd)	
end if

function UpdCookie(IdProd, QtdProd)
	dim sTxt
	
	sTxt = mid(request.cookies("Okidata"),instr(request.cookies("Okidata"), IdProd),len(mid(request.cookies("Okidata"),1,instr(request.cookies("Okidata"), IdProd))
	
	if instr(request.cookies("Okidata"), sTxt) = 0 then
		'response.cookies("Okidata") = replace(request.cookies("Okidata"), sTxt, "")
		'response.cookies("Okidata") = IdProd & "," & QtdProd & "@" & request.cookies("Okidata")		
	end if
end function
%>