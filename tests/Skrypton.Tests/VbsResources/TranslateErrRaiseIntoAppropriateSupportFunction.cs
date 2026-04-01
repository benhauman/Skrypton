
            _.RAISEERROR(VBScriptConstants.vbObjectError);
            _.RAISEERROR(VBScriptConstants.vbObjectError, "Source");
            _.RAISEERROR(VBScriptConstants.vbObjectError, "Source", "Test");
            _.CALLm1v4(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "RAISEERROR", VBScriptConstants.vbObjectError, "Source", "Test", "Bonus Argument");
            _.CLEARANYERROR();
            _.CLEARANYERROR();
            _.CALLm1v1(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "CLEARANYERROR", "Bonus Argument");
