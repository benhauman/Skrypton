            if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v5(this, _env.hlObj ?? throw new InvalidOperationException("Reference not set:"), "GetValue", "ServiceRequestRecordSpecific.TargetDateScheduled", (Int16)0, (Int16)0, (Int16)0, (Int16)0)), (Int16)1)))
            {
                _.SETm1a0(this, _env.ComboBoxPriority ?? throw new InvalidOperationException("Reference not set:"), "Disabled", "true");
            }
            else
            {
                _.SETm1a0(this, _env.ComboBoxPriority ?? throw new InvalidOperationException("Reference not set:"), "Disabled", "false");
            }
