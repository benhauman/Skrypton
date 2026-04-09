Dim serv
Set serv = Nothing
IF not serv is nothing THEN
  IF serv.enabled(7) = True THEN
    serv.Enabled(8) = False
  ELSE
    serv.Enabled(9) = True
  END IF
END IF
