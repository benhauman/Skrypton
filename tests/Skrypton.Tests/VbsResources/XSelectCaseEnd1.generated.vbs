Dim status, internalstatus
status = "z"
SELECT CASE status
  CASE "SRMInternalStateOpen"
    internalstatus = "OPEN"
  CASE "SRMInternalStateToBeChecked"
    internalstatus = "TOBECHECKED"
  CASE "SRMInternalStateWaitingForAnalyse"
    internalstatus = "WAITING_FOR"
  CASE "SRMInternalStateWaitingForOffer"
    internalstatus = "WAITING_FOR"
  CASE "SRMInternalStateOrderReceived"
    internalstatus = "TOBECHECKED"
  CASE "SRMInternalStateSolved"
    internalstatus = "SOLVED"
  CASE "SRMInternalStateClosed"
    internalstatus = "CLOSED"
  CASE "SRMInternalStateWaitingForInternal"
    internalstatus = "WAITING_FOR"
  CASE ELSE
    internalstatus = "OPEN"
END SELECT
