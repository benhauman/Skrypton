Dim parmval, adoSQLCmdParam
WITH adoSQLCmdParam
  .Pr.Ap  .CreateParameterX("RETURN_VALUEx", 3, 4)
  .Parameters.Append  .CreateParameterY("@FirstCharName", 202, 1, 1, "FirstCharName")
  .Execute
  parmval =.Parameters(2).Value
END WITH
