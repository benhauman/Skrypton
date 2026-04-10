
            _.RAISEERROR(VBScriptConstants.vbObjectError);
            _.RAISEERROR(VBScriptConstants.vbObjectError, "Source");
            _.RAISEERROR(VBScriptConstants.vbObjectError, "Source", "Test");
            _.RAISEERROR(VBScriptConstants.vbObjectError, "Source", "Test", "Bonus Argument");
            _.CLEARANYERROR();
            _.CLEARANYERROR();
            _.CALLm1v1(this, _, "CLEARANYERROR", "Bonus Argument");
