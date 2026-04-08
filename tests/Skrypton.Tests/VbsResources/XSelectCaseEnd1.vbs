Dim status, internalstatus
status = "z"
Select Case status 
	Case "SRMInternalStateOpen" internalstatus = "OPEN"
	Case "SRMInternalStateToBeChecked" internalstatus= "TOBECHECKED"
	Case "SRMInternalStateWaitingForAnalyse" internalstatus= "WAITING_FOR"
	Case "SRMInternalStateWaitingForOffer" internalstatus = "WAITING_FOR"
	Case "SRMInternalStateOrderReceived" internalstatus = "TOBECHECKED"
	Case "SRMInternalStateSolved" internalstatus = "SOLVED"
	Case "SRMInternalStateClosed" internalstatus = "CLOSED"
	Case Else internalstatus = "OPEN"
	Case "SRMInternalStateWaitingForInternal" internalstatus= "WAITING_FOR"
End Select