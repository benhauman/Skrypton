Public Sub TestFso()
  task.SetValue "RoutingHelper.AgentID", 0, 0, 0, Person.GetValue("HLOBJECTINFO.ID", 0, 0, task.GetSvcUnitCount(), 0)
End Sub
