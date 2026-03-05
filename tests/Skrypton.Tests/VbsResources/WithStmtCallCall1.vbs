	dim parmval, adoSQLCmdParam
	With adoSQLCmdParam
	 Set .ActiveConnection = Nothing
	.Pr.Ap .CreateParameterX("RETURN_VALUEx", 3, &H0004 )
	.Parameters.Append .CreateParameterY("@FirstCharName", _
	202, &H0001,1,"FirstCharName") 
	.Execute
	parmval = .Parameters(2).Value
	End With
