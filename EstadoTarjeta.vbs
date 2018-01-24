'##########################################################################################################################
'####
'####  Fecha de creación:		07/11/2017
'####  Autor: 					Jhonny Montoya – TCS.
'####  Descripción: 			El componente modfica el Estado de las tarjetas para provar el estado bloqueado y los restablece en el segundo llamado
'####  Parámetros:				sResPasoAnt: Párametro de entrada que sirve para evaluar el resultado del componente anterior.
'####  ________________________________________________
'####  Fecha de modificación: 	[dd/mm/yyyy]
'####  Modificado por: 			[nombre y apellido] – [nombre de la empresa del proveedor].
'####  Descripción cambio: 		[Descripción del script, función o subrutina].
'####VDIB0053 
'##########################################################################################################################

Option explicit
On error resume next

'Declaración de variables locales
Dim rs,sQuery,sError, nNumRowsUpdated
Dim StrEstadoTarje, blEstadoViejo, StrEstado

If parameter("sResPasoAnt")="000" Then '1
	If DDT("NumTarjeta") <> "" and DDT("NumCuenta") <> "" Then'2
		If (DDT("Escenario")="TarjInacIseries") Then	'3
			Set rs = CreateObject("ADODB.Recordset") 
			If oConexion.State=0 Then
				oConexion.Open(DDTParam("sConnectionIseries"))	
			End If

			'Query para modificar Estado
			err.number = 0
			
			'================================================================================
			If DDT.Exists("EstadoTarjeta") Then 
			    StrEstadoTarje = DDT("EstadoTarjeta")
				blEstadoViejo = 0
			Else
	
				sQuery = "Select TJESTTARDB from <<CABFFTARJ>> where TJNROTRJ =  " & Right (DDT("NumTarjeta"),10)
				sQuery = Replace(sQuery,"<<CABFFTARJ>>", DDI("CABFFTARJ") )				
				
				rs.Open sQuery, oConexion
				If err.number <> 0 Then
					sError = sError & "Error al capturar el estado de la cuenta original: " & err.Description
				End If
				
				If rs.EOF = False and rs.BOF = False Then
					StrEstado = rs.Fields("TJESTTARDB").Value
					
				Else
					Reporter.ReportEvent micPass,"Estado Tarjeta", "No se encontró el estado de la Tarjeta" & sQuery
					parameter("sResPaso") = "000"
				End if
				DDT.Add "EstadoTarjeta",StrEstado	
				blEstadoViejo = 1
				End If
				
						
			'Query para modificar Estado
			err.number = 0
						
			If blEstadoViejo = 0 Then
				sQuery = "Update <<CABFFTARJ>> set TJESTTARDB = '" & DDT("EstadoTarjeta") & "' WHERE TJNROTRJ =" & Right(DDT("NumTarjeta"),10) 
				StrEstado = "A"
			Else
				sQuery = "Update <<CABFFTARJ>> set TJESTTARDB = 'L' WHERE TJNROTRJ =" & Right(DDT("NumTarjeta"),10)
				StrEstado = "L"
			End if
			sQuery = Replace(sQuery,"<<CABFFTARJ>>", DDI("CABFFTARJ") )				
					
			oConexion.Execute sQuery,nNumRowsUpdated
			If err.number <> 0 Then
				parameter("sResPaso") = "001"
				sError = sError & "Error al actualizar el estado de la tarjeta: " & err.Description
			Else  
				Reporter.ReportEvent micPass,"Estado Tarjeta", "Número de la Tarjeta: " & DDT("NumTarjeta") + vbCrLf + "Estado de la tarjeta Final: " & StrEstado
				parameter("sResPaso") = "000"
			End If

			rs.Close
			
			If sError <> "" Then
				Reporter.ReportEvent micFail, "Configurar Parametros", "No fue posible modificar algunos parametros: " & sError	
				Parameter("sResPaso") = "001"
			End If
			
		Else
			Reporter.ReportEvent micFail,"Escenario no esta", "El escenario no esta definido en el componente"	
			Parameter("sResPaso") = "001"
		End If'3	
	Else
		Reporter.ReportEvent micFail,"Configurar Parametros", "La cuenta o la tarjeta estan en blanco"	
		Parameter("sResPaso") = "001"
	End If'2
Else
	Reporter.ReportEvent micFail,"Configurar Parametros", "No es posible ejecutar componente porque componente anterior falló."	
	parameter("sResPaso") = "001"	
End If'1
